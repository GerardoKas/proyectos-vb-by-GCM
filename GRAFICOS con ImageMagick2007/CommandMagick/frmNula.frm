VERSION 5.00
Object = "{6446818E-FC0F-4E35-9697-05A0AAB43FD0}#19.0#0"; "GCMPicControl.ocx"
Begin VB.Form Form1 
   Caption         =   "Command Magick By GCM 2005"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7140
   Icon            =   "frmNula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMiniatura 
      Caption         =   "Crear Miniatura al Cargar y al Probar efectos"
      Height          =   240
      Left            =   1980
      TabIndex        =   26
      Top             =   315
      Value           =   1  'Checked
      Width           =   3525
   End
   Begin VB.Frame Frame2 
      Height          =   825
      Left            =   1980
      TabIndex        =   20
      Top             =   5310
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox txtDestino 
         Height          =   285
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Efecto Destino"
         ToolTipText     =   "Efecto Destino"
         Top             =   495
         Width           =   1995
      End
      Begin VB.TextBox txtOrigen 
         Height          =   285
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Efecto Origen"
         ToolTipText     =   "Efecto de Partida"
         Top             =   495
         Width           =   2040
      End
      Begin VB.TextBox txtComando 
         Height          =   285
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "COMANDO"
         Top             =   180
         Width           =   1680
      End
      Begin VB.TextBox txtEfActual 
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0"
         ToolTipText     =   "Num Efecto Actual"
         Top             =   180
         Width           =   420
      End
      Begin VB.TextBox txtViendo 
         Height          =   285
         Left            =   2205
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Vid"
         ToolTipText     =   "Viendo? (1,0)"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdDeshacer 
      Caption         =   "Quitar Ultimo"
      Enabled         =   0   'False
      Height          =   285
      Left            =   855
      TabIndex        =   17
      Top             =   4320
      Width           =   1050
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Añadir"
      Enabled         =   0   'False
      Height          =   285
      Left            =   45
      TabIndex        =   16
      Top             =   4320
      Width           =   780
   End
   Begin VB.CommandButton cmdApplyCopia 
      Caption         =   "Aplicar a Copia"
      Enabled         =   0   'False
      Height          =   330
      Left            =   45
      TabIndex        =   12
      Top             =   5850
      Width           =   1500
   End
   Begin VB.TextBox txtOriginal 
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "FILENAME ORIGINAL"
      ToolTipText     =   "Fichero Original Cargado"
      Top             =   0
      Width           =   5055
   End
   Begin PictureControl_GCM.GCM_Picture gcm1 
      Height          =   4785
      Left            =   1980
      TabIndex        =   10
      Top             =   495
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   8440
   End
   Begin VB.ListBox lstEfectos 
      Height          =   840
      Left            =   0
      TabIndex        =   9
      Top             =   4635
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   0
      TabIndex        =   2
      Top             =   1485
      Width           =   1860
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   285
         Left            =   765
         TabIndex        =   19
         Top             =   2295
         Width           =   825
      End
      Begin VB.CommandButton cmdProbar 
         Caption         =   "Ver"
         Enabled         =   0   'False
         Height          =   285
         Left            =   225
         TabIndex        =   18
         Top             =   2295
         Width           =   510
      End
      Begin VB.VScrollBar vscroll 
         Height          =   1590
         Index           =   2
         Left            =   1395
         TabIndex        =   8
         Top             =   315
         Width           =   240
      End
      Begin VB.VScrollBar vscroll 
         Height          =   1590
         Index           =   1
         Left            =   765
         TabIndex        =   7
         Top             =   315
         Width           =   240
      End
      Begin VB.VScrollBar vscroll 
         Height          =   1590
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   240
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   15
         Top             =   90
         Width           =   90
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Index           =   1
         Left            =   855
         TabIndex        =   14
         Top             =   90
         Width           =   90
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   90
         Width           =   90
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   2
         Left            =   1350
         TabIndex        =   5
         Top             =   1935
         Width           =   315
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Index           =   1
         Left            =   945
         TabIndex        =   4
         Top             =   1935
         Width           =   90
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   1935
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdApplyAll 
      Caption         =   "Aplicar Al Original"
      Enabled         =   0   'False
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   5490
      Width           =   1500
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "frmNula.frx":038A
      Left            =   0
      List            =   "frmNula.frx":03BE
      TabIndex        =   0
      Top             =   0
      Width           =   1905
   End
   Begin VB.Frame frmChangeEfect 
      Height          =   1455
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
      Begin VB.CommandButton cmdChange 
         Caption         =   "Cambiar Efecto"
         Height          =   555
         Left            =   90
         TabIndex        =   28
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "EFECTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   405
         TabIndex        =   29
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnSaveList 
         Caption         =   "Guardar Lista Efectos"
      End
      Begin VB.Menu mnuLoadList 
         Caption         =   "Cargar Lista Efectos"
      End
      Begin VB.Menu mnuProbarList 
         Caption         =   "Probar Lista Efectos"
      End
      Begin VB.Menu mnuClearEfects 
         Caption         =   "Borrar Efectos Añadidos"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Datos para los scrolls de efectos
Private sFuncion As String
Private sParams As String
Private TipoScroll As String
Private hayScrolls As Integer
Private values(2) As String
Private vMult(2) As Double
Private vMax(2) As Long
Private vMin(2) As Long
Private vStep(2) As Double
Private vAddScroll(2) As Long
'Efecto anterior por si ya fue aplicado
Private prevComando As String

'Datos del Ini
Const IniFile = "efectosMagick.ini"
Dim loadedIni As Boolean
Dim colEfectos As New Collection

'===========================================
Private Sub cmdChange_Click()
List1.Visible = True
frmChangeEfect.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'MsgBox KeyAscii
If KeyAscii = 104 Then
    If Frame2.Visible = True Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
End If
End Sub

Private Sub Form_Load()
Set colEfectos = New Collection
mostrarScroll -1
loadIniFile
End Sub

Private Sub Form_Resize()
On Error Resume Next
gcm1.Width = Me.ScaleWidth - gcm1.Left
gcm1.Height = Me.ScaleHeight - gcm1.Top
End Sub

Private Sub gcm1_onLoadImage(filename As String, cancel As Boolean)
'si es solo un efecto, continuar
verLista True
verProbar True

If gcm1.external = False Then Exit Sub
waitAll
myFile = filename
txtOriginal = myFile
'el numefecto es MUY importante. Indica que se abrio unaq nueva imagen
txtEfActual = 0 'reseteamos los efectos a 0
'creamos el efecto0 = miniatura
If chkMiniatura.value = 0 Then
ImagesEfectos(0) = myFile
Else
ImagesEfectos(0) = getEfectFileName(myFile, 0) 'pedimos un nombre de efecto
doFileMini 325, myFile, ImagesEfectos(0)
End If
filename = ImagesEfectos(0)
resumeAll
End Sub

Private Sub cmdProbar_Click()
'variables en esta rutina
'txtEfActual
'txtComando
'txtOrigen
'txtDestino
If gcm1.mypicture = "" Then MsgBox "No has cargado imagen": Exit Sub
If txtEfActual > maxEfectos Then MsgBox "no se pueden aplicar mas efectos": Exit Sub
If txtComando = prevComando Then MsgBox "Ya estas viendo ese efecto": Exit Sub
txtViendo = 1
waitAll
'el anterior efecto o 0
If txtEfActual = 0 Then txtEfActual = 1
txtOrigen = ImagesEfectos(txtEfActual - 1)
txtDestino = getEfectFileName(myFile, txtEfActual)
'ejecutar comando
If executeMagick(txtComando, txtOrigen, txtDestino) = False Then
    MsgBox "Error Mio"
    GoTo terminar
End If
gcm1.external = False
gcm1.LoadNew txtDestino
'ver botones
terminar:
verLista False
verProbar True
verAplicar False
verAgregar True
resumeAll
End Sub

Private Sub cmdCancelar_Click()
Dim ef As String
verLista True
verAgregar False
verProbar True
verAplicar True
'cmdAdd.Enabled = False
'cmdDeshacer.Enabled = True
'cmdApplyAll.Enabled = True
'cmdApplyCopia.Enabled = True
'cmdCancelar.Enabled = False
'Frame1.Visible = False
If txtEfActual = 0 Then Exit Sub
txtViendo = ""
ef = ImagesEfectos(txtEfActual - 1)
If ef = gcm1.mypicture Then Exit Sub
gcm1.LoadNew ef
End Sub

Private Sub cmdDeshacer_Click()
If txtEfActual = 0 Then Exit Sub
If lstEfectos.ListCount = 0 Then Exit Sub

txtEfActual = txtEfActual - 1
lstEfectos.RemoveItem lstEfectos.ListCount - 1
prevComando = ""
gcm1.LoadNew ImagesEfectos(txtEfActual)
End Sub

Private Sub cmdAdd_Click()
lstEfectos.AddItem txtComando
List1_Click
txtEfActual = txtEfActual + 1
txtViendo = ""
verLista True
verAgregar False
verAplicar True
End Sub

Private Sub cmdApplyAll_Click()
waitAll
BackupIt myFile, myFile & ".BAK"
For i = 0 To lstEfectos.ListCount - 1
    executeMagick lstEfectos.list(i), myFile, myFile
Next
'setear todo a cero pq la original ya cambio
lstEfectos.Clear
txtEfActual = 0
ImagesEfectos(txtEfActual) = myFile

resumeAll
End Sub

Private Sub cmdApplyCopia_Click()
waitAll
mycopy$ = duplyIt(myFile)
For i = 0 To lstEfectos.ListCount - 1
    executeMagick lstEfectos.list(i), mycopy, mycopy
Next
MsgBox ("Guardado en :" & vbCrLf & mycopy)
resumeAll

End Sub

Private Sub List1_Click()
lblDesc(0) = "?"
lblDesc(1) = "?"
lblDesc(2) = "?"
TipoScroll = ""
lblName.Caption = List1.Text
If loadedIni = False Then
    EfectosOriginales
Else
    loadefectoini (List1.Text)
End If

verLista False
verAplicar True
verProbar True

End Sub
'=== LISTA DE EFECTOS Y COMANDOS PARA REAPLICAR
Private Sub lstEfectos_Click()
PopupMenu mnMenu
End Sub

Private Sub mnSaveList_Click()
frmEfectFiles.loadfile = False
frmEfectFiles.Show 1, Me
End Sub

Private Sub mnuClearEfects_Click()
lstEfectos.Clear
End Sub

Private Sub mnuLoadList_Click()
'Load frmEfectFiles
frmEfectFiles.loadfile = True
frmEfectFiles.Show 1, Me
End Sub

Private Sub mnuProbarList_Click()
txtEfActual = 0
For i = 0 To lstEfectos.ListCount - 1
    txtComando = lstEfectos.list(i)
    cmdProbar_Click
    MsgBox "Hecho el " & txtEfActual
    txtEfActual = txtEfActual + 1
    Next
End Sub
'==== TXTS RAROS
Private Sub txtDestino_Change()
ImagesEfectos(txtEfActual) = txtDestino
End Sub

Private Sub txtViendo_Change()
If txtViendo = "1" Then
    List1.Visible = False
    prevComando = txtComando
Else
    List1.Visible = True
    prevComando = ""
End If
End Sub
'============================================0
'==========SCROLLLLLSSSSSSS ==================

Private Sub vscroll_Change(index As Integer)
values(index) = getScrollValue(index)
lbl1(index).Caption = values(index)
Select Case LCase(TipoScroll)
    Case "normal"
        sParams = values(0)
        For i = 1 To hayScrolls
            sParams = sParams & "," & values(i)
        Next
    Case "percent"
    'si hay que añadir el pcentaje
        sParams = values(0) & "%"
        For i = 1 To hayScrolls
            sParams = sParams & "," & values(i) & "%"
        Next
    Case Else
    'si es repetir la funcion +++ como en -contrast
       sParams = ""
        For i = 1 To values(0) - 1
            sParams = sFuncion & " " & sParams
        Next
End Select
txtComando = sFuncion & " " & sParams
End Sub

Private Sub mostrarScroll(uboundScrolls)
    For i = 0 To uboundScrolls
        vscroll(i).Visible = True
        lbl1(i).Visible = True
        lblDesc(i).Visible = True
        'values(i) = 0
    Next
    For i = uboundScrolls + 1 To 2
        vscroll(i).Visible = False
        lbl1(i).Visible = False
        lblDesc(i).Visible = False
    Next
    hayScrolls = uboundScrolls
    'If hayScrolls = -1 Then TipoScroll = "Plus"
    sParams = ""
End Sub

Private Function isSingle(value)
isSingle = IIf(CInt(value) <> value, True, False)
End Function

Private Function getScrollValue(index)
X = vscroll(index).value
VALOR$ = (Round(X * (vMult(index)) + vAddScroll(index), 2))
VALOR = Replace(VALOR, ",", ".")
getScrollValue = VALOR
End Function

Private Sub valoresScroll(max, min, init, step, Optional index)
    'numPasos = (max - min) / step
    'multiscroll = pasos
    m = 1
    i = 0
    Dim multscroll As Double
    Dim sumscroll As Long
    
    Do While isSingle(max) Or isSingle(min) Or isSingle(step)
        i = i + 1
        m = 10 * i
        max = max * m
        min = min * m
        step = step * m
        init = init * m
        If m > 10000 Then
            MsgBox "esos valores no funcionan :" & min & " " & max & " " & step & " ...ERROR"
        End If
    Loop
    multscroll = 1 / m
    
    If min < 0 Then
        sumscroll = min
        min = min - min
        max = max + min
        init = init + min
    End If
    large = Int(max / 10) + 1 'diez pasos en large change
        
        'convertir valores a entero
    If IsMissing(index) Then
        For i = 0 To hayScrolls
            vMult(i) = multscroll
            vAddScroll(i) = sumscroll
            vscroll(i).max = min
            vscroll(i).min = max
            vscroll(i).value = init
            vscroll(i).LargeChange = Abs(large)
            vscroll(i).SmallChange = step
        Next
    Else
        vMult(index) = multscroll
        vAddScroll(index) = sumscroll
        vscroll(index).max = min
        vscroll(index).min = max
        vscroll(index).value = init
        vscroll(index).LargeChange = Abs(large)
        vscroll(index).SmallChange = step
    End If
End Sub



'////////////////////////////////////////////////////////////////////
Sub waitAll()
Me.MousePointer = 11
End Sub
Sub resumeAll()
Me.MousePointer = 0
End Sub

Sub verAgregar(ver As Boolean)
If ver = True Then
    cmdAdd.Enabled = True
    cmdDeshacer.Enabled = True
Else
    cmdAdd.Enabled = False
    cmdDeshacer.Enabled = False
    If List1.ListCount > 0 Then
        cmdDeshacer.Enabled = True
    End If
End If
End Sub
Sub verAplicar(ver As Boolean)
If ver = True Then
    cmdApplyAll.Enabled = True
    cmdApplyCopia.Enabled = True
Else
    cmdApplyAll.Enabled = False
    cmdApplyCopia.Enabled = False
End If
End Sub
Sub verProbar(ver As Boolean)
If ver = True Then
    cmdProbar.Enabled = True
    cmdCancelar.Enabled = True
Else
    cmdProbar.Enabled = False
    cmdCancelar.Enabled = False
End If
End Sub
Sub verLista(ver As Boolean)
If ver = True Then
    List1.Visible = True
    List1.Enabled = True
    frmChangeEfect.Visible = False
    Frame1.Visible = False
Else
    List1.Visible = False
    frmChangeEfect.Visible = True
    Frame1.Visible = True
End If
End Sub

Sub EfectosOriginales()
Select Case List1.Text
    Case "Gamma"
        sFuncion = "-gamma"
        'CallByName Me, sFuncion, VbMethod
        lblDesc(0) = "G.Rojo"
        lblDesc(1) = "G.Verde"
        lblDesc(2) = "G.Azul"
        TipoScroll = "Normal"
        mostrarScroll 2
        valoresScroll 4, 0, 1, 0.1
    Case "AutoNormalize"
        sFuncion = "-normalize"
        mostrarScroll -1
        valoresScroll 1, 1, 1, 1
    Case "AutoContrast"
        sFuncion = "-contrast"
        mostrarScroll 0
        valoresScroll 10, 1, 1, 1
    Case "AutoEqualize"
        sFuncion = "-equalize"
        mostrarScroll -1
        valoresScroll 1, 1, 1, 1
    Case "AutoDestramado"
        sFuncion = "-despeckle"
        mostrarScroll 0
        valoresScroll 10, 1, 1, 1
    Case "Perfilado"
        sFuncion = "-unsharp"
        lblDesc(0) = "Radio"
        lblDesc(1) = "Sigma"
        TipoScroll = "Normal"
        mostrarScroll 1
        valoresScroll 50, 0, 0, 0.5
    Case "Difuminado"
        sFuncion = "-blur"
        lblDesc(0) = "Radio"
        lblDesc(1) = "Sigma"
        TipoScroll = "Normal"
        mostrarScroll 1
        valoresScroll 50, 0, 0, 0.5
    Case "Modulate"
        'modulate bright contrast an hue
        sFuncion = "-modulate"
        lblDesc(0) = "Brillo"
        lblDesc(1) = "Saturacion"
        lblDesc(2) = "Tonalidad"
        TipoScroll = "Normal"
        mostrarScroll 2
        valoresScroll 255, 0, 100, 2
    Case "ReducirRuido"
        sFuncion = "-noise"
        lblDesc(0) = "Radio"
        TipoScroll = "Normal"
        mostrarScroll 0
        valoresScroll 50, 0, 0, 1
    Case "B/W Level"
        sFuncion = "-level"
        lblDesc(0) = "Negro"
        lblDesc(1) = "Blanco"
        lblDesc(2) = "Gamma"
        TipoScroll = "Normal"
        mostrarScroll 1
        valoresScroll 100, 0, 50, 1
    Case "Swirl"
        sFuncion = "-swirl"
        lblDesc(0) = "grados"
        TipoScroll = "Normal"
        mostrarScroll 1
        valoresScroll 720, -720, 0, 1
    Case "Roll"
        sFuncion = "-roll"
        lblDesc(0) = "X"
        lblDesc(1) = "Y"
        mostrarScroll 1
        TipoScroll = "Normal"
        valoresScroll 500, -500, 0, 1
    Case "Solarize"
        sFuncion = "-solarize"
        mostrarScroll 0
        TipoScroll = "Normal"
        valoresScroll 100, 0, 0, 1
    Case "Rotate"
        sFuncion = "-rotate"
        mostrarScroll 0
        TipoScroll = "Normal"
        valoresScroll 360, 0, 0, 1
    Case "Implode"
        sFuncion = "-implode"
        mostrarScroll 0
        TipoScroll = "Normal"
        valoresScroll 100, 0, 0, 0.5
    Case "Explode"
        sFuncion = "+implode"
        mostrarScroll 0
        TipoScroll = "Normal"
        valoresScroll -100, 0, 0, 0.5
    Case "Median"
        sFuncion = "-median"
        mostrarScroll 0
        TipoScroll = "Normal"
        valoresScroll 9, 0, 0, 1
    Case "Colors"
        sFuncion = "+dither -colors"
        mostrarScroll 0
        TipoScroll = "Normal"
        valoresScroll 255, 0, 64, 1
    Case Else
'        MsgBox "Me olvidé"
End Select
End Sub


'==========FICHERO INI CON EFECTOS ==============

Sub loadIniFile()
If Dir$(IniFile, vbNormal) <> "" Then
    If FileLen(IniFile) > 0 Then
    setIniFile IniFile
    loadEfectos
    loadedIni = True
    End If
Else
    loadedIni = False
End If
End Sub

Sub loadEfectos()
List1.Clear
strsecciones = ReadSections()
lssecciones = Split(strsecciones, Chr(0))
For Each i In lssecciones
    If i <> "" Then
        Desc = ReadFromFile(CStr(i), "name")
        colEfectos.Add i, Desc
        List1.AddItem Desc
    End If
Next
End Sub

Sub loadefectoini(name As String)
efid$ = colEfectos(name)
numscrolls = ReadFromFile(efid, "numScrolls")
hayScrolls = numscrolls - 1 'pq empieza en 0
equal = ReadFromFile(efid, "Equal")
' LCase(equal) = "true"
mostrarScroll hayScrolls
sFuncion = ReadFromFile(efid, "funcion")
txtComando = sFuncion
TipoScroll = ReadFromFile(efid, "tipo")
If numscrolls > 0 Then
    For i = 0 To hayScrolls
    If LCase(equal) = "true" Or numscrolls = 1 Then
        it = "": salir = True
    Else
        it = i: salir = False
    End If
    max = ReadFromFile(efid, "Max" & it)
    min = ReadFromFile(efid, "Min" & it)
    init = ReadFromFile(efid, "Init" & it)
    step = Val(ReadFromFile(efid, "Step" & it))
    If salir Then
        valoresScroll max, min, init, step
        Exit For
    Else
        valoresScroll max, min, init, step, i
    End If
    Next
'====
End If

End Sub

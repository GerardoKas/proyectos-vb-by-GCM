VERSION 5.00
Begin VB.UserControl GCM_Picture 
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   KeyPreview      =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   4440
   Begin VB.Frame frmcosa 
      Height          =   1770
      Left            =   1305
      TabIndex        =   8
      Top             =   855
      Visible         =   0   'False
      Width           =   2310
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   765
         TabIndex        =   11
         Top             =   1395
         Visible         =   0   'False
         Width           =   870
         Begin VB.CommandButton cmdBack 
            Caption         =   "<"
            Height          =   285
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   420
         End
         Begin VB.CommandButton cmdForward 
            Caption         =   ">"
            Height          =   285
            Left            =   450
            TabIndex        =   12
            Top             =   0
            Width           =   420
         End
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   90
         Pattern         =   "*.JPG"
         TabIndex        =   9
         Top             =   180
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4140
      TabIndex        =   3
      Top             =   3105
      Width           =   285
      Begin VB.Label lblMenos 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lblMas 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   45
         Width           =   90
      End
   End
   Begin VB.VScrollBar scrollY 
      Height          =   3090
      LargeChange     =   25
      Left            =   4140
      Max             =   25
      SmallChange     =   10
      TabIndex        =   2
      Top             =   0
      Width           =   240
   End
   Begin VB.HScrollBar scrollX 
      Height          =   240
      LargeChange     =   25
      Left            =   0
      Max             =   25
      SmallChange     =   10
      TabIndex        =   1
      Top             =   3075
      Width           =   4140
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      Height          =   2520
      Left            =   0
      MousePointer    =   2  'Cross
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   2460
      ScaleWidth      =   3540
      TabIndex        =   0
      Top             =   0
      Width           =   3600
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   240
         Left            =   405
         TabIndex        =   10
         Top             =   1350
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "MARKS"
         Height          =   195
         Left            =   2925
         TabIndex        =   5
         Top             =   270
         Width           =   570
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         DrawMode        =   7  'Invert
         Height          =   375
         Left            =   855
         Top             =   45
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "SIZES"
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   225
         Width           =   465
      End
   End
End
Attribute VB_Name = "GCM_Picture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public mypicture As String

Private imageWidth As Long
Private imageHeight As Long
Public zoom As Single
Public imagePixelWidth As Long
Public imagePixelHeight As Long
Public doDebug As Boolean
Public external As Boolean
Private Cargado As Boolean
Private pw As Single, ph As Single
Private posx As Single
Private posy As Single
Private oPic As IPictureDisp
Private Const divScroll As Single = 25 '25 + 1.756 = 26.46 ... jo...
Private Const porScroll As Double = 1.756
Private Const twipsPicture As Double = 26.46
Private Const TwipsScreen As Single = 15
Private marcando As Boolean
'Private marcaIni As String
'Private marcaFin As String

Event changedZoom(zoom As Single, pixelsWidth As Long, pixelsHeight As Long)
Event changedPos(posx, posy)
Event changedMark(x, y, w, h)
Event zIniciadaMarka(x, y)
Event zArrastrando(w, h)
Event newFileLoaded(filename)
Event onLoadImage(filename As String, cancel As Boolean)

Event onForward(nowFile As String)
Event onBackward(nowFile As String)

Private markX As Long
Private markY As Long
Private markW As Long
Private markH As Long
Public marcaLeft As Long
Public marcaTop As Long
Public marcaWidth As Long
Public marcaHeight As Long

Const addZoom = 0.25

Public Sub changePos(posx, posy)
If Cargado = False Then Exit Sub
scrollX.value = posx
scrollY.value = posy
End Sub


Private Sub lblMas_Click()
If Cargado = False Then Exit Sub
changeZoom (zoom + addZoom)

End Sub

Private Sub lblMenos_Click()
If Cargado = False Then Exit Sub
changeZoom (zoom - addZoom)
End Sub

Private Sub Picture1_GotFocus()
scrollX.Visible = False
scrollY.Visible = False

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'x e y estan en screentwips
'el desplazamiento de la imagen se mide por el scroll ..
'que hay que pasar a imagenTwips
If Button = 1 And mypicture <> "" Then
Dim mx As Double, my As Double
Shape1.Move x, y
mx = (picTwips(getScroll(scrollX)) * porScroll) + screenTwips(x)
my = (picTwips(getScroll(scrollY)) * porScroll) + screenTwips(y)
markX = Round(mx / zoom)
markY = Round(my / zoom)
Label2.Caption = vbCrLf & markX & " x " & markY
RaiseEvent zIniciadaMarka(markX, markY)
marcando = True
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim w As Long, h As Long
With Shape1
If marcando = True And Button = 1 Then
w = x - .left
h = y - .top
If w < 0 Then w = 0
If h < 0 Then h = 0
.width = w
.height = h
markW = Round(screenTwips(w) / zoom)
markH = Round(screenTwips(h) / zoom)
Label2.Caption = marcaIni & vbCrLf & " Size: " & markW & ", " & markH
RaiseEvent zArrastrando(markW, markH)
End If
End With
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
copyMarks
RaiseEvent changedMark(markX, markY, markW, markH)
marcando = False
End If
End Sub

Private Sub copyMarks()
marcaLeft = markX
marcaTop = markY
marcaWidth = markW
marcaHeight = markH
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.Files.Count >= 1 Then
external = True
LoadNew Data.Files(1)
End If
End Sub

Private Sub scrollX_Change()
RaiseEvent changedPos(scrollX, scrollY)
refrescar
End Sub

Private Sub scrollY_Change()
RaiseEvent changedPos(scrollX, scrollY)
refrescar
End Sub

Public Sub LoadNew(fichero As String)
Dim cancel As Boolean
RaiseEvent onLoadImage(fichero, cancel)
DoEvents
If cancel = True Then Exit Sub
mypicture = fichero
If mypicture = "" Then MsgBox "No Se Especifico Fichero": Exit Sub
If Dir$(mypicture, vbNormal) = "" Then MsgBox "NO EXISTE:" & vbCrLf & mypicture, vbOKOnly + vbExclamation, "GCMPICTURE ERROR": Exit Sub
On Error GoTo Err1
Set oPic = LoadPicture(mypicture)
imageWidth = oPic.width
imageHeight = oPic.height
Cargado = True
changePos 0, 0
If zoom = 0 Then zoom = 1
changeZoom zoom
RaiseEvent newFileLoaded(mypicture)
external = False
On Error Resume Next
getFilesInDir (fichero)
Exit Sub
Err1:
MsgBox "ESE FICHERO PARECE NO EXISTIR" & vbCrLf & Err.Number & " : " & Err.Description, vbQuestion, "GCMPICTURE ERROR "
End Sub


Public Sub refrescar()
If Not Cargado Then Exit Sub
posx = getScroll(scrollX)
posy = getScroll(scrollY)
Label1.Caption = "TWIPS : " & vbCrLf
Label1.Caption = Label1.Caption & "imageW=" & imageWidth & " - imageH=" & imageHeight & vbCrLf
Label1.Caption = Label1.Caption & "controlW=" & Picture1.ScaleWidth & " - controlH:" & Picture1.ScaleHeight & vbCrLf
Label1.Caption = Label1.Caption & "zumTw_W=" & pw & " - zumTw_H=" & ph & vbCrLf
Label1.Caption = Label1.Caption & vbCrLf & "PIXELS : "
Label1.Caption = Label1.Caption & vbCrLf & "RealW:" & picTwips(imageWidth) & " - RealH:" & picTwips(imageHeight)
Label1.Caption = Label1.Caption & vbCrLf & "ZumW:" & picTwips(pw) & " - ZumH:" & picTwips(ph) & vbCrLf
Label1.Caption = Label1.Caption & vbCrLf & "SCROLLS : "
Label1.Caption = Label1.Caption & vbCrLf & "scrollMaxX:" & getScroll(scrollX.Max) & " - scrollMaxY:" & getScroll(scrollY.Max) & vbCrLf
Label1.Caption = Label1.Caption & "PosX:" & (posx) & " - PosY:" & (posy) & vbCrLf
Label1.Caption = Label1.Caption & "Zoom:" & zoom
Picture1.ToolTipText = mypicture
imagePixelWidth = picTwips(imageWidth)
imagePixelHeight = picTwips(imageHeight)
If doDebug = False Then
    Label1.Visible = False
    Label2.Visible = False
Else
    Label1.Visible = True
    Label2.Visible = True
End If
Call Picture1_Paint

End Sub

Private Sub Picture1_Paint()
If oPic Is Nothing Then Exit Sub
Picture1.Cls
Picture1.PaintPicture oPic, -posx, -posy, (pw), (ph), 0, 0, imageWidth, imageHeight
End Sub

Public Sub zoomAjustar()
Dim myzX As Single, myzY As Single, fz As Single
If Cargado = False Then Exit Sub
myzX = Picture1.ScaleWidth / imageWidth
myzY = Picture1.ScaleHeight / imageHeight
fz = IIf(myzX < myzY, myzX, myzY) * porScroll
If fz < 0.15 Then fz = 0.15
changeZoom fz
End Sub

Public Sub changeZoom(myzoom As Single)
If Not Cargado Then Exit Sub
If myzoom < 0.15 Then Exit Sub
If myzoom > 6 Then Exit Sub
zoom = myzoom
pw = imageWidth * zoom
ph = imageHeight * zoom
RaiseEvent changedZoom(zoom, picTwips(pw), picTwips(ph))

On Error Resume Next
scrollX.Max = ((pw / porScroll) - Picture1.width) / divScroll
scrollY.Max = ((ph / porScroll) - Picture1.height) / divScroll
If scrollX.Max < 1 Then
scrollX.Max = 0
Else
scrollX.LargeChange = Int(scrollX.Max / 10)
End If
If scrollY.Max < 1 Then
scrollY.Max = 0
Else
scrollY.LargeChange = Int(scrollY.Max / 10)
End If
refrescar
End Sub

Private Sub UserControl_Initialize()
doDebug = False
dibujarLineas
Shape1.Move 0, 0, 1, 1
End Sub

Private Sub UserControl_GotFocus()
scrollX.Visible = True
scrollY.Visible = True
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
Picture1.left = 0
Picture1.top = 0
Picture1.width = ScaleWidth ' - scrollY.Width
Picture1.height = ScaleHeight ' - scrollX.Height
scrollX.top = Picture1.height - scrollX.height
scrollY.left = Picture1.width - scrollY.width
scrollX.width = Picture1.width
scrollY.height = Picture1.height
If Not Cargado Then
    dibujarLineas
End If
Frame1.left = scrollX.width
Frame1.top = scrollY.height
End Sub

Private Sub dibujarLineas()
Picture1.Cls
Picture1.Align = center
Picture1.FontSize = 16
Picture1.Print "DROP AN IMAGE HERE" & vbCrLf & "SUELTA UNA IMAGEN AQUI"
Picture1.Line (0, 0)-(Picture1.ScaleWidth, Picture1.ScaleHeight)
Picture1.Line (Picture1.width, 0)-(0, Picture1.height)
'Picture1.FontSize = 28
End Sub

'Public Sub Marcar(x As Long, y As Long, w As Long, h As Long)
'Shape1.Move picTwips(x), picTwips(y), picTwips(w), picTwips(h)
'End Sub

'picturePixel2Twips
Public Function picTwips(value)
'la medida que usa el picture para sus handles. no es himetric. no se que es
picTwips = Round(value / twipsPicture)
End Function

'screenPixel2Twips
Public Function screenTwips(value)
'twipsperpixel es = 15
screenTwips = Round(value / TwipsScreen)
End Function

Private Function getScroll(value As Long)
getScroll = value * divScroll
End Function

Private Function getFilesInDir(file)
Dim nF As String, p As Long, myDir As String
On Error GoTo otroErr
p = InStrRev(file, "\")
myDir = left(file, p - 1)
fname = Right(file, Len(file) - p)
'nF = Dir$(myDir & "\*.jpg", vbNormal)
File1.Pattern = "*.jpg;*.jpeg;*.gif;*.bmp;*.png;*.ico"
File1.Path = myDir
'buscar pos actual
For i = 0 To File1.ListCount - 1
   If File1.List(i) = fname Then
      File1.ListIndex = i
      Exit For
   End If
Next

Exit Function
otroErr:
MsgBox "ocurrio un error al leer el directorio"
End Function

Private Sub cmdBack_Click()
If File1.ListIndex = 0 Then mensaje ("Primera Imagen"): Exit Sub
File1.ListIndex = File1.ListIndex - 1
external = True
LoadNew File1.Path & "\" & File1.List(File1.ListIndex)
End Sub

Private Sub cmdForward_Click()
If File1.ListIndex = File1.ListCount - 1 Then mensaje ("Fin del Directorio"): Exit Sub
File1.ListIndex = File1.ListIndex + 1
external = True
LoadNew File1.Path & "\" & File1.List(File1.ListIndex)
End Sub



Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode
Select Case KeyCode
Case 34 'av pag
    v = scrollY.value + scrollY.LargeChange
    scrollY.value = IIf(v < scrollY.Max, v, scrollY.Max)
Case 33 're pag
    v = (scrollY.value - scrollY.LargeChange)
    scrollY.value = IIf(v > 0, v, 0)
Case 36 'home
    v = scrollX.value - scrollX.LargeChange
    scrollX.value = IIf(v > 0, v, 0)
    
Case 35 ' fin
    v = scrollX.value + scrollX.LargeChange
    scrollX.value = IIf(v < scrollX.Max, v, scrollX.Max)
Case Else
    'MsgBox KeyCode
End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
Select Case (KeyAscii)
Case Asc("j")
    zoomAjustar
Case Asc("+")
    lblMas_Click
Case Asc("-")
    lblMenos_Click
Case Asc(" ")
    cmdForward_Click
Case (8)
    cmdBack_Click
Case Else
    UserControl_GotFocus
    'MsgBox "Que deberia hacer la Tecla " & Chr(KeyAscii) & " ?" & vbCrLf & "Prueba : 'j' '+' '-'"
End Select
KeyAscii = 0
End Sub

Public Sub ChangeMark(left As Long, top As Long, width As Long, height As Long)
markX = left
markY = top
markW = width
markH = height
copyMarks
moveRectangleToMarks
RaiseEvent changedMark(markX, markY, markW, markH)
End Sub

Private Sub moveRectangleToMarks()
'cambiar pixeles a imagetwips
'las markas estan en pixeles
'zoomAjustar

Shape1.left = (screenPixel2Twips(markX) - (scrollX * divScroll)) * zoom
Shape1.top = (screenPixel2Twips(markY) - (scrollY * divScroll)) * zoom
Shape1.width = (screenPixel2Twips(markW) - (scrollX * divScroll)) * zoom
Shape1.height = (screenPixel2Twips(markH) - (scrollY * divScroll)) * zoom

End Sub

Private Function picPixel2Twips(pixels)
picPixel2Twips = Round(pixels * twipsPicture)
End Function

Private Function screenPixel2Twips(pixels)
screenPixel2Twips = Round(pixels * TwipsScreen)
End Function

Private Sub Command1_Click()
ChangeMark 100, 100, 500, 300
End Sub

Private Sub mensaje(texto)
MsgBox texto, vbOKOnly, "GCM_PICTURE"

End Sub

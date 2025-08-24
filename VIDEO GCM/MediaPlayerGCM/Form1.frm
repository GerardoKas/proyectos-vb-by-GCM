VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   Caption         =   "VISOR DE MULTIMEDIA CON MEDIAPLAYER2"
   ClientHeight    =   5085
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   6975
   DrawMode        =   7  'Invert
   FillColor       =   &H00000040&
   FillStyle       =   2  'Horizontal Line
   ForeColor       =   &H0000C0C0&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0442
   MousePointer    =   4  'Icon
   ScaleHeight     =   5085
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FFFF&
      Height          =   645
      ItemData        =   "Form1.frx":074C
      Left            =   90
      List            =   "Form1.frx":074E
      MouseIcon       =   "Form1.frx":0750
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   240
      Width           =   6840
   End
   Begin VB.Label lblTip 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Searching ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3060
      TabIndex        =   3
      Top             =   1935
      Visible         =   0   'False
      Width           =   3300
   End
   Begin MediaPlayerCtl.MediaPlayer MP1 
      Height          =   3375
      Left            =   45
      TabIndex        =   0
      Top             =   1665
      Width           =   6855
      AudioStream     =   -1
      AutoSize        =   -1  'True
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   4210688
      DisplayForeColor=   65535
      DisplayMode     =   0
      DisplaySize     =   0
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label props 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Duracion Nombre ..."
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   1425
   End
   Begin VB.Menu zFind 
      Caption         =   "BUSCAR"
      Begin VB.Menu mnuSearch 
         Caption         =   "SEARCH..."
      End
      Begin VB.Menu mnFavi 
         Caption         =   "VIDEO"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnFmp3 
         Caption         =   "MP3"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnFWAV 
         Caption         =   "WAV"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnFMIdi 
         Caption         =   "MIDI"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu zLista 
      Caption         =   "LISTA"
      Begin VB.Menu mnBorrarLista 
         Caption         =   "BORRAR"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSaveList 
         Caption         =   "GUARDAR"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnLista 
         Caption         =   "LISTA LARGA"
      End
   End
   Begin VB.Menu zopciones 
      Caption         =   "Opciones"
      Begin VB.Menu mnAuto 
         Caption         =   "AUTOPLAY"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnFull 
         Caption         =   "FULESCRIN"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnstats 
         Caption         =   "ESTADISTICAS"
      End
      Begin VB.Menu mnRate 
         Caption         =   "FRECUENCIA (RATE)"
      End
      Begin VB.Menu mnControles 
         Caption         =   "OCULTAR CONTROLES"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unidad As String
Dim fileahora As String
Const defaultList = "List-GCM-Player.txt"

Private Sub list1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim c As Integer
If (Data.GetFormat(vbCFFiles)) Then

For i = 1 To Data.Files.Count
    ListaArk2(NumArks) = mkArkivo(Data.Files.Item(i))
    List1.AddItem (ListaArk2(NumArks).nomArk)
    NumArks = NumArks + 1
Next
End If
On Error Resume Next
MP1.FileName = (Data.Files(i))
MP1.Open (Data.Files(i))

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
App.TaskVisible = False
'MsgBox (KeyCode)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Savelist (App.path & "\" & defaultList)
Unload FrmRate
Unload frmList
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> vbMinimized Then
List1.Top = 0 'props.Height
List1.Left = 0
List1.Height = 700
List1.Width = Me.ScaleWidth
MP1.Left = 0
MP1.Width = Me.ScaleWidth ' - List1.Width
MP1.Height = Me.ScaleHeight - List1.Height
MP1.Top = List1.Height
'List1.Height = Me.ScaleHeight - List1.Top - 255
'HScroll1.Top = Me.ScaleHeight - 255
'lblrate.Top = HScroll1.Top
End If
End Sub
Private Sub Form_Load()
'DoEvents
If loadlist(defaultList) Then
    AñadirArkivos
    Load frmList '.AñadirList
End If
unidad = Busca.LaUNIDAD(App.path)
'MP1.TransparentAtStart = True

MP1.ShowStatusBar = True
'##ESTO ES OPCIONAL :
'BUSCAR AVIS EN EL DISCO
Me.Visible = True
    'mnFavi_Click
    'mnFmp3_Click
'mnAuto_Click
    'TOCAR EL PRIMERO
'    If List1.ListCount > 0 Then
'    List1.ListIndex = 0
'    List1_DblClick
'    End If
'##-----------------

End Sub

Private Sub List1_DblClick()
If (List1.SelCount <> 0) Then
ACTUAL = List1.ListIndex
fileahora = getFilename(ListaArk2(List1.ListIndex))
Tocar (fileahora)
End If
End Sub

Private Sub mnAuto_Click()
If mnAuto.Checked = True Then
mnAuto.Checked = False
'MP1.AutoStart = False
Else
mnAuto.Checked = True
'MP1.AutoStart = True
End If

End Sub

'
'

Private Sub mnBorrarLista_Click()
clearList
frmList.clearList
NumArks = 0
End Sub

'Private Sub mncodecs_Click()
'Dim cnt As Integer
'cnt = MP1.CodecCount
'For i = 0 To cnt
'cad = cad & Chr(13) & MP1.GetCodecDescription(1)
'Next
'cad = MP1.GetMediaInfoString(mpClipAuthor)
'MsgBox (cad)
'End Sub

Private Sub mnControles_Click()
If (mnControles.Caption = "OCULTAR CONTROLES") Then
mnControles.Caption = "VER CONTROLES"
MP1.ShowControls = False
Else
mnControles.Caption = "OCULTAR CONTROLES"
MP1.ShowControls = True
End If
End Sub

Private Sub mnFull_Click()
MP1.DisplaySize = mpFullScreen
End Sub

Private Sub mnLista_Click()
'frmList.Show 0, Me
frmList.Visible = True
frmList.ZOrder 0
End Sub

Private Sub mnRate_Click()
Set FrmRate.Elmp = MP1
FrmRate.Show 1, Me
End Sub

Private Sub mnstats_Click()
'MP1.ShowDialog mpShowDialogContextMenu
MP1.ShowDialog mpShowDialogStatistics
End Sub

Private Sub mnuSaveList_Click()
Savelist defaultList
End Sub

Private Sub mnuSearch_Click()
frmSearch.Show 1, Me
End Sub

'SIGUE CON EL SIGUIENTE ARCHIVO
Private Sub MP1_EndOfStream(ByVal Result As Long)
    If List1.ListCount <> 0 Then
        ACTUAL = ACTUAL + 1
        If (ACTUAL >= List1.ListCount) Then ACTUAL = 0
        List1.ListIndex = ACTUAL
        frmList.marcar (ACTUAL)
        Tocar getFilename(ListaArk2(ACTUAL))
    End If
End Sub



'Private Sub MP1_OpenStateChange(ByVal OldState As Long, ByVal NewState As Long)
''MsgBox (OldState)
''MP1.Play
''If NewState = mpReadyStateComplete Then
'If NewState < 6 Then
''Mostrar
'End If
'
''Me.Caption = "old open: " & OldState & "NUEVO OPEN: " & NewState
''MOstrar
''End If
'End Sub

'Private Sub MP1_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
'Me.Caption = ListaArk2(ACTUAL).nomArk & " ... " & newPosition
'End Sub


Public Sub Tocar(Archivo As String)
Dim oldVol As Long
On Error GoTo unerr
'oldVol = MP1.get
MP1.Cancel
MP1.AutoSize = False
MP1.DisplaySize = mpFitToSize

'MP1.AutoStart = mnAuto.Checked
MP1.FileName = Archivo
MP1.Play
'MP1.ShowDisplay = True

Do
DoEvents
Loop While MP1.ReadyState < 4
Mostrar
MP1.Volume = oldVol

Exit Sub
unerr:
Elerror ("ALTOCAR")
End Sub

Public Function getSecs(Archivo As String) As String
On Error GoTo unerr
MP1.Cancel
MP1.AutoStart = False
MP1.Open Archivo
Do
DoEvents
Loop While MP1.ReadyState < 4
getSecs = (MP1.Duration)
Exit Function
'MP1.Stop
MP1.Cancel
unerr:
Elerror ("ALCARGARLENGTH")
End Function

Private Sub Elerror(cuando As String)
    If MP1.ErrorCode <> 0 Then MsgBox MP1.ErrorDescription & Chr(13) & cuando, , "ERROR DEL ACTIVEX ..."
    'If Err.Number <> 0 Then MsgBox Err.Description & Chr(13) & cuando, , "ERROR DE ESTE PROGRAMA ..."
End Sub
Private Sub Mostrar()
Dim filesize
On Error GoTo unerr
If MP1.Duration = 0 Then Exit Sub
props.Caption = ""
props.Caption = props.Caption & "DURA: " & SecsToTime(MP1.Duration) & Chr(13)
If MP1.ImageSourceWidth <> 0 Then
props.Caption = props.Caption & "WxH: " & MP1.ImageSourceWidth & "x" & MP1.ImageSourceHeight & Chr(13)
End If
filesize = FileLen(fileahora)
props.Caption = props.Caption & "SIZE:" & TheSize(FileLen(fileahora))

Exit Sub
unerr:
Elerror "MOSTRANDO PARAMETROS"
End Sub

Public Sub AñadirArkivos()
ACTUAL = 0
For i = 0 To NumArks - 1
List1.AddItem ListaArk2(i).nomArk
Next
End Sub

Public Sub clearList()
List1.Clear
End Sub

Sub lockForm()
Me.MousePointer = vbHourglass
lblTip.Visible = True
Me.Enabled = False
End Sub

Sub unlockForm()
Me.MousePointer = vbDefault
lblTip.Visible = False
Me.Enabled = True
End Sub

Private Sub MP1_ScriptCommand(ByVal scType As String, ByVal Param As String)
Debug.Print scType
End Sub

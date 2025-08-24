VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Magick A Mano"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLog 
      Caption         =   "Log"
      Height          =   285
      Left            =   4095
      TabIndex        =   24
      Top             =   2700
      Width           =   600
   End
   Begin VB.CommandButton cmdAddToLog 
      Caption         =   "AddToLog"
      Height          =   285
      Left            =   3105
      TabIndex        =   23
      Top             =   2700
      Width           =   960
   End
   Begin VB.TextBox txtCMD 
      BackColor       =   &H80000004&
      Height          =   645
      Left            =   45
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   3240
      Width           =   4830
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "INF"
      Height          =   330
      Left            =   4545
      TabIndex        =   21
      Top             =   810
      Width           =   420
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   3840
      Left            =   4950
      TabIndex        =   16
      Top             =   0
      Width           =   5100
      ExtentX         =   8996
      ExtentY         =   6773
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
   Begin VB.Frame frmContains 
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   360
      Width           =   2445
      Begin VB.CommandButton cmdClearListFiles 
         Caption         =   "Borrar"
         Height          =   285
         Left            =   3960
         TabIndex        =   15
         Top             =   1260
         Width           =   735
      End
      Begin VB.CommandButton cmdListToText 
         Caption         =   "Crear Listsa"
         Height          =   285
         Left            =   2700
         TabIndex        =   13
         Top             =   1260
         Width           =   1140
      End
      Begin VB.ListBox lstFicheros 
         Columns         =   1
         Height          =   840
         Left            =   2700
         TabIndex        =   12
         Top             =   315
         Width           =   2040
      End
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   90
         Pattern         =   "*.JPG;*.gif;*.png;*.bmp;*.ico"
         TabIndex        =   9
         Top             =   315
         Width           =   2445
      End
      Begin VB.Label Label4 
         Caption         =   "scalewidth * 1950"
         Height          =   465
         Left            =   225
         TabIndex        =   14
         Top             =   1845
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Add Fichero(s) a la lista "
         Height          =   240
         Left            =   90
         TabIndex        =   11
         Top             =   90
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-->"
         Height          =   195
         Left            =   1125
         TabIndex        =   10
         Top             =   1305
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "EXEC"
      Height          =   465
      Left            =   990
      TabIndex        =   5
      Top             =   2700
      Width           =   1275
   End
   Begin VB.TextBox txtComando 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0FF&
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Text            =   "%CMD-N% %ORIG-n% %DEST%"
      Top             =   2340
      Width           =   4335
   End
   Begin VB.CommandButton cmdRevolver 
      Caption         =   "Pasar A original"
      Height          =   330
      Left            =   3375
      TabIndex        =   3
      Top             =   1485
      Width           =   1365
   End
   Begin VB.ComboBox cmbComandos 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "COMANDOS"
      Top             =   1935
      Width           =   4380
   End
   Begin VB.CommandButton cmdFolder 
      Caption         =   "Directorio"
      Height          =   285
      Left            =   3330
      TabIndex        =   1
      Top             =   45
      Width           =   1005
   End
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Text            =   "c:\"
      Top             =   45
      Width           =   3165
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "PREV"
      Height          =   465
      Left            =   45
      TabIndex        =   6
      Top             =   2700
      Width           =   915
   End
   Begin VB.CommandButton cmdPlusName 
      Caption         =   "+"
      Height          =   330
      Left            =   45
      TabIndex        =   17
      Top             =   1485
      Width           =   285
   End
   Begin VB.CommandButton cmdAddOriginal 
      Caption         =   "Add"
      Height          =   330
      Left            =   3330
      TabIndex        =   18
      Top             =   810
      Width           =   510
   End
   Begin VB.CommandButton cmdDeleteOriginal 
      Caption         =   "Delete"
      Height          =   330
      Left            =   3870
      TabIndex        =   19
      Top             =   810
      Width           =   645
   End
   Begin VB.TextBox txtDestino 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   405
      TabIndex        =   20
      Text            =   "DESTINO"
      Top             =   1485
      Width           =   2850
   End
   Begin VB.ListBox lstFilesOriginal 
      Height          =   645
      Left            =   90
      OLEDropMode     =   1  'Manual
      TabIndex        =   25
      Top             =   720
      Width           =   3120
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "No Error"
      Height          =   195
      Left            =   2295
      TabIndex        =   7
      Top             =   2925
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hadError As Boolean
Dim lastFile As String
Dim numEfs As Integer
Dim txtFiles As String
Const imgLine = "<a href=""%FILE%"">%FILE%</a> - <img align=middle src=""%FILE%"" width=100><br>"

Private Sub lstFilesOriginal_Click()
showFile lstFilesOriginal.List(lstFilesOriginal.ListIndex)
End Sub

Private Sub lstFilesOriginal_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
fname = Data.files(1)
Set filesys = CreateObject("Scripting.FileSystemObject")
odir = filesys.getparentfoldername(fname)
oname = filesys.getbasename(fname)
ext = filesys.getextensionname(fname)
txtFolder = odir
lstFilesOriginal.AddItem """" & oname & "." & ext & """"
getNewDestName oname & "." & "jpg"
End Sub

Private Sub cmdAddOriginal_Click()
t = InputBox("Introduce File o comando", , "'xc:white'")
lstFilesOriginal.AddItem t
End Sub

Private Sub cmdAddToLog_Click()
Static logfile As String
Static nowDir As String
If nowDir <> txtFolder Or logfile = "" Then
    'pedir filename
    fname = InputBox("Nombre de Archivo de Log", "Guardar Log", txtFolder & "Comandos.txt")
    If fname = "" Then MsgBox "No se guardara log": Exit Sub
    logfile = fname
End If
f = FreeFile
Open logfile For Append As #f
Print #f, txtComando
Close #f
lblResult.Caption = "Added To Log"
End Sub

Private Sub cmdClearListFiles_Click()
lstFicheros.Clear
End Sub

Private Sub cmdDeleteOriginal_Click()
If lstFilesOriginal.ListIndex > -1 Then
lstFilesOriginal.RemoveItem (lstFilesOriginal.ListIndex)
End If
End Sub

Private Sub cmdExecute_Click()
Dim cmd As String
Call cmdPrev_Click
cmd = txtComando.Text
If chkLog.Value = 1 Then
    cmdAddToLog_Click
End If

Dim WshShell As Object
Screen.MousePointer = vbHourglass

dopause = " & echo. & echo Error: %errorlevel% & pause"
executecmd = "cmd /C " & Left(txtFolder, 2) & " & cd """ & txtFolder & """ & echo INI & convert.exe -monitor " & cmd & " & echo FIN" & dopause

txtCMD.Text = executecmd

Set WshShell = CreateObject("WScript.Shell")
r = WshShell.Run(executecmd, 5, True)

lblResult = "Error:" & r
Screen.MousePointer = vbArrow
If r <= 0 Then addComando cmbComandos.Text

showFile txtDestino.Text
End Sub

Function addComando(texto As String)
For i = 0 To cmbComandos.ListCount - 1
    If cmbComandos.List(i) = texto Then
        Exit Function
    End If
Next
cmbComandos.AddItem texto
End Function

Private Sub cmdInfo_Click()
infoFile lstFilesOriginal.List(lstFilesOriginal.ListIndex)
End Sub

Private Sub cmdListToText_Click()
'cmbFilesOriginal.Clear
For i = 0 To lstFicheros.ListCount - 1
    temp = lstFicheros.List(i)
    temp = escapeName(temp)
    lstFilesOriginal.AddItem """" & temp & """"
Next
'cmbFilesOriginal.Text = lstFicheros.ListCount & " Ficheros Origen"
getNewDestName lstFicheros.List(0)
frmContains.Width = Label1.Width
frmContains.Height = Label1.Height
End Sub

Private Sub cmdPrev_Click()
Dim files As String
Dim cmd As String
Dim dest As String
files = getFilesText
cmd = cmbComandos.Text
dest = """" & txtDestino.Text & """"
txtComando = cmd & " " & files & " " & dest
Me.Caption = "Viewing Cmd"
End Sub

Private Sub cmdRevolver_Click()
Call Renew
End Sub

Private Sub File1_Click()
showFile File1.filename
End Sub

Private Sub File1_DblClick()
lstFicheros.AddItem File1.filename
End Sub


Private Sub Form_Load()
txtFolder = GetSetting("M", "A", "folder", "c:\")
On Error Resume Next
File1.Path = txtFolder
loadData cmbComandos, "usedCommands.txt"
configWeb
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "M", "A", "folder", txtFolder
saveData cmbComandos, "usedCommands.txt"
End Sub

Private Sub cmdFolder_Click()
txtFolder = checkfolder(txtFolder)
On Error Resume Next
File1.Path = txtFolder
End Sub

Function checkfolder(fld As String)
If Right(fld, 1) <> "\" Then
    fld = fld & "\"
End If
checkfolder = fld
End Function

Private Sub Label1_Click()
frmContains.Width = 4850
frmContains.Height = 1750
End Sub

Sub Renew()
lstFilesOriginal.Clear
lstFilesOriginal.AddItem txtDestino.Text
getNewDestName txtDestino.Text
End Sub

Sub getNewDestName(Optional fname As String)
Dim rex As RegExp, matches As MatchCollection
If fname = "" Then fname = txtDestino.Text
If fname = "" Then fname = lastFile
Set filesys = CreateObject("Scripting.FileSystemObject")
oname = filesys.getbasename(fname)
ext = filesys.getextensionname(fname)
odir = filesys.getparentfoldername(fname)
Set rex = New RegExp
rex.IgnoreCase = True
rex.Pattern = "\d+$"
If rex.Test(oname) = True Then
    Set matches = rex.Execute(oname)
    oname = rex.Replace(oname, matches(0) + 1)
Else
    oname = oname & "-F0"
End If
txtDestino.Text = oname & "." & ext
End Sub

Private Sub cmdPlusName_Click()
getNewDestName
End Sub

Sub configWeb()
f = FreeFile()
Open "vacio.html" For Output As #f
Print #f, "<html><body>0 IMAGENES</body></html>"
Close #f
web1.Navigate App.Path & "\vacio.html"
Do While web1.Busy = True
    DoEvents
Loop
web1.Document.body.leftMargin = 0
web1.Document.body.topMargin = 0
End Sub

Function getFilesText()
txtFiles = ""
For i = 0 To lstFilesOriginal.ListCount - 1
    txtFiles = txtFiles & " " & lstFilesOriginal.List(i)
Next
getFilesText = txtFiles
End Function


Function escapeName(oname)
oname = Replace(oname, "[", "\[")
oname = Replace(oname, "]", "\]")
oname = Replace(oname, """", "")
escapeName = oname
End Function


Sub saveData(combo As ComboBox, filename As String)
f = FreeFile
Close
Open filename For Output As #f
For i = 0 To combo.ListCount - 1
    Print #f, combo.List(i)
Next
Close #f
End Sub

Sub loadData(combo As ComboBox, filename As String)
f = FreeFile
Open filename For Input As #f
Do While Loc(f) <> LOF(f)
    Line Input #f, v
    combo.AddItem v
Loop
Close #f
End Sub

Function showFile(file As String)
Dim it As String, doc As Object
Set doc = web1.Document
file = mkPath(txtFolder, file)
it = Replace(imgLine, "%FILE%", file)
doc.open
doc.write (it)
'doc.Close
End Function

Function infoFile(file As String)
Dim it As String, doc As Object
Set doc = web1.Document
absfile = mkPath(txtFolder, file)
txtfile = "infoF.txt"
Set WshShell = CreateObject("WScript.Shell")
r = WshShell.Run("cmd /c identify -verbose """ & absfile & """>" & txtfile, 5, True)
If r > 0 Then MsgBox "Error Al Leer " & file: Exit Function
f = FreeFile()
txt = ""
Open txtfile For Input As #f
Do While Loc(f) <> LOF(f) And Not EOF(f)
    Line Input #f, Lin
    txt = txt & "<br>" & Lin
Loop
Close #f
it = Replace(imgLine, "%FILE%", absfile)
allText = it & "<hr>" & txt
doc.open
doc.write (allText)
End Function

Function isFile(cadena As String)
Set filesys = CreateObject("Scripting.FileSystemObject")
isFile = filesys.fileExists(cadena)
End Function

Function mkPath(ruta As String, file As String)
Set filesys = CreateObject("Scripting.FileSystemObject")
mkPath = filesys.buildPath(ruta, escapeName(file))
End Function

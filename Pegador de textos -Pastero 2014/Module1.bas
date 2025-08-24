Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_DESKTOP = 0
Public Const HWND_TOP = 0

Public grabado As Boolean
Public Const maxW = 450
Public Const maxH = 400
Public Const maxArea = 1600
Public Const maxTitle = 32
Public Const minw = 200
Public Const minh = 100

Public availW As Long
Public availH As Long

Public esMaxi As Boolean
Public savedir As String
Public filename As String
Public savedFilename As String
Public fileWasOpened As Boolean
Public Const bar = vbCrLf & "(...)" & vbCrLf

Public Const MyApp = "GCM_Pastero_Drop_n_Save"
Public Const MySection = "Config"

Function ponerArriba(frmHwnd As Long, Optional x = 200, Optional y = 0)
Dim hwnd As Long
hwnd = frmHwnd
'w = Int(Screen.Width / Screen.TwipsPerPixelX)

ok = SetWindowPos(hwnd, -1, x, y, minw, minh, 0)
End Function

Function sizeMini()
If Form1.WindowState = vbMaximized Then
 Form1.WindowState = vbNormal
End If
Form1.ScaleMode = vbPixels
'X = Form1.ScaleLeft
'Y = Form1.ScaleWidth
Form1.Width = minw * Screen.TwipsPerPixelX
Form1.Height = minh * Screen.TwipsPerPixelY
End Function

Function sizeMaxi()
'X = Form1.ScaleLeft
'Y = Form1.ScaleTop
On Error Resume Next
Form1.Width = maxW * Screen.TwipsPerPixelX
Form1.Height = maxH * Screen.TwipsPerPixelY
End Function

Function getFilename(dialog As CommonDialog)
On Error GoTo ErrName
dialog.Flags = cdlOFNNoValidate
dialog.CancelError = True
dialog.Filter = "Texto plano *.txt|*.txt|Pagina Web *.html|*.html|Todos los Archivos *.*|*.*"
dialog.DefaultExt = "txt"
If savedir <> "" Then
    dialog.InitDir = savedir
End If
dialog.filename = filename
dialog.ShowSave
getFilename = dialog.filename

Exit Function
ErrName:
If Err.Number = 32755 Then
    getFilename = ""
End If
End Function


Function grabarFichero(fichero As String)
f = FreeFile
If fichero = "" Then Exit Function
Open fichero For Binary As #f
Put #f, 1, Form1.Text1.Text
Close #f
grabado = True
End Function

Function antesDeBorrar()
If grabado = True Then antesDeBorrar = True: Exit Function

resp = MsgBox("No se han guardado Los cambios, deseas Guardarlo?", vbYesNoCancel + vbApplicationModal + vbQuestion + vbMsgBoxSetForeground, "Grabar?")
If resp = vbYes Then
    nf = getFilename(Form1.cmdlg)
    If nf = "" Then
        antesDeBorrar = True
        Exit Function
    End If
    grabarFichero (nf)
    antesDeBorrar = True
ElseIf resp = vbCancel Then
    antesDeBorrar = False
Else
    antesDeBorrar = True
End If
End Function

Function crearTitulo(texto As String)
Dim tit As String
'comprobar puntuacion no posible
If Form1.Text1.Tag = "Nodo" Then Exit Function
If fileWasOpened = True Then
    crearTitulo = basename(savedFilename)
    Exit Function
End If
tit = Left(texto, maxTitle)
tit = limpiar(tit)
p = InStr(1, tit, Chr(13), vbBinaryCompare)
If (p > 0) Then
tit = Left(tit, p - 1)
End If
tit = changeBadChars(tit)

filename = tit & ".txt"
crearTitulo = filename
End Function

Function limpiar(texto As String)
'sacar enteres al principio y al final. por si el texto viene despues
Dim cad As String
cad = texto
While (Left(cad, 2) = vbCrLf)
cad = Right(cad, Len(cad) - 2)
Wend
While (Right(cad, 2) = vbCrLf)
cad = Left(cad, Len(cad) - 2)
Wend
cad = Trim(cad)
limpiar = cad
End Function


Function changeBadChars(texto As String)
texto = Replace(texto, ":", "-")
texto = Replace(texto, "?", "!")
texto = Replace(texto, """", "'")
texto = Replace(texto, "/", "-")
texto = Replace(texto, "\", "-")
texto = Replace(texto, "|", "-")
texto = Replace(texto, "<", "[")
texto = Replace(texto, ">", "]")
texto = Replace(texto, "*", "·")
changeBadChars = texto
End Function

'Set objFld = obj.BrowseForFolder(hwnd, "Titulo Ventana", Bif_Flags, CarpetaOrigen)

Public Function BrowseForFolder(folder As String)
Dim obj, objFld, objPath
Dim objOrigen
'Flags of Style for the Popup
Const BIF_NEWDIALOGSTYLE = &H40
Const BIF_RETURNONLYFSDIRS = &H1     ' For finding a folder to start document searching
Const BIF_DONTGOBELOWDOMAIN = &H2    ' For starting the Find Computer
Const BIF_STATUSTEXT = &H4
Const BIF_RETURNFSANCESTORS = &H8
Const BIF_EDITBOX = &H10
Const BIF_VALIDATE = &H20             ' insist on valid result (or CANCEL)
Const BIF_BROWSEFORCOMPUTER = &H1000   ' Browsing for Computers.
Const BIF_BROWSEFORPRINTER = &H2000    ' Browsing for Printers
Const BIF_BROWSEINCLUDEFILES = &H4000  ' Browsing for Everything
'Common Folders For Open
Const ssfALTSTARTUP = &H1D
Const ssfAPPDATA = &H1A
Const ssfBITBUCKET = &HA
Const ssfCOMMONALTSTARTUP = &H1E
Const ssfCOMMONAPPDATA = &H23
Const ssfCOMMONDESKTOPDIR = &H19
Const ssfCOMMONFAVORITES = &H1F
Const ssfCOMMONPROGRAMS = &H17
Const ssfCOMMONSTARTMENU = &H16
Const ssfCOMMONSTARTUP = &H18
Const ssfCONTROLS = &H3
Const ssfCOOKIES = &H21
Const ssfDESKTOP = &H0
Const ssfDESKTOPDIRECTORY = &H10
Const ssfDRIVES = &H11
Const ssfFAVORITES = &H6
Const ssfFONTS = &H14
Const ssfHISTORY = &H22
Const ssfINTERNETCACHE = &H20
Const ssfLOCALAPPDATA = &H1C
Const ssfMYPICTURES = &H27
Const ssfNETHOOD = &H13
Const ssfNETWORK = &H12
Const ssfPERSONAL = &H5
Const ssfPRINTERS = &H4
Const ssfPRINTHOOD = &H1B
Const ssfPROFILE = &H28
Const ssfPROGRAMFILES = &H26
Const ssfPROGRAMS = &H2
Const ssfRECENT = &H8
Const ssfSENDTO = &H9
Const ssfSTARTMENU = &HB
Const ssfSTARTUP = &H7
Const ssfSYSTEM = &H25
Const ssfTEMPLATES = &H15
Const ssfWINDOWS = &H24

'On Error GoTo noSel
Set obj = CreateObject("Shell.Application")
Set objOrigen = obj.NameSpace(folder)
Set objFld = obj.BrowseForFolder(Form1.hwnd, "Buscando Carpeta :", BIF_NEWDIALOGSTYLE Or BIF_RETURNONLYFSDIRS, "")
If objFld Is Nothing Then
BrowseForFolder = ""
Else
Set objPath = objFld.Items.Item
BrowseForFolder = objPath.Path
End If

End Function

Public Function openFile(fichero As String)
Dim contenido As String
Dim tamano As Long
Dim f As Integer
Dim pos2 As Integer
f = FreeFile
If Left(fichero, 1) = """" Then
    pos2 = InStrRev(fichero, """")
    fichero = Mid(fichero, 2, pos2 - 2)
End If
savedFilename = fichero
fileWasOpened = True

savedir = basepath(fichero)
Open fichero For Binary As #f
tamano = FileLen(fichero)
contenido = Input(tamano, #f)
Close #f
contenido = Replace(contenido, Chr(0), Chr(12))
contenido = replaceVblf(contenido)
openFile = contenido
End Function

Public Function replaceVblf(texto) As String
Dim reg As RegExp
Set reg = New RegExp
reg.Global = True
reg.MultiLine = True
reg.Pattern = "\n"
replaceVblf = reg.Replace(texto, vbCrLf)
End Function

Public Function basepath(file)
Dim pos As Integer
    pos = InStrRev(file, "\")
    basepath = Left(file, pos)
End Function

Public Function basename(file)
    Dim pos As Integer
    pos = InStrRev(file, "\")
    basename = Right(file, Len(file) - pos)
End Function

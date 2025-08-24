Attribute VB_Name = "Module1"
Option Explicit


Public FileList As Collection
Public sDestino As String
Public meWidth As Long
Public meHeight As Long
'----------------------------------------------------------
Public Const myApp = "GcmDownloader"
Public Const mySection = "Params"
Public Const myCaption = "Downloader By Gcm 2005"
'---------------------------------------------------
Public Const msgOk = "ok" 'mensaje de ok para buscar los correctos
Public Const msgSep = " - " ' Separador del mensaje en el listbox

'------------------------------------------------------------
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_DESKTOP = 0
Public Const HWND_TOP = 0
'-----------------------------------------------------------

Public Declare Function URLDownloadToFile _
Lib "urlmon" _
Alias "URLDownloadToFileA" ( _
ByVal pCaller As Long, _
ByVal szURL As String, _
ByVal szFileName As String, _
ByVal dwReserved As Long, _
ByVal lpfnCB As Long) As Long


Public Enum URLDownloadErrConstants
ErrSuccess = &H0
ErrErrorUnknown = &H1
ErrAborted = &H80004004
ErrDestFileExists = &H800C0001
ErrInvalidUrl = &H800C0002
ErrNoSession = &H800C0003
ErrCannotConnect = &H800C0004
ErrResourceNotFound = &H800C0005
ErrObjectNotFound = &H800C0006
ErrDataNotAvailable = &H800C0007
ErrDownloadFailure = &H800C0008
ErrAuthenticationRequired = &H800C0009
ErrNoValidMedia = &H800C000A
ErrConnectionTimeout = &H800C000B
ErrInvalidRequest = &H800C000C
ErrUnknownProtocol = &H800C000D
ErrSecurityProblem = &H800C000E
ErrCannotLoadData = &H800C000F
ErrCannotInstantiateObject = &H800C0010
ErrRedirectFailed = &H800C0014
ErrRedirectToDir = &H800C0015
ErrCannotLockRequest = &H800C0016
End Enum

Public URLDownloadError As URLDownloadErrConstants
Public URLDownloadErrorDescription As String

Public Function GetDownloadErrorDescription()

Select Case URLDownloadError
Case ErrSuccess
URLDownloadErrorDescription = "ErrSuccess"
Case ErrErrorUnknown
URLDownloadErrorDescription = "ErrErrorUnknown"
Case ErrAborted
URLDownloadErrorDescription = "ErrAborted"
Case ErrDestFileExists
URLDownloadErrorDescription = "ErrDestFileExists"
Case ErrInvalidUrl
URLDownloadErrorDescription = "ErrInvalidUrl"
Case ErrNoSession
URLDownloadErrorDescription = "ErrNoSession"
Case ErrCannotConnect
URLDownloadErrorDescription = "ErrCannotConnect"
Case ErrResourceNotFound
URLDownloadErrorDescription = "ErrResourceNotFound"
Case ErrObjectNotFound
URLDownloadErrorDescription = "ErrObjectNotFound"
Case ErrDataNotAvailable
URLDownloadErrorDescription = "ErrDataNotAvailable"
Case ErrDownloadFailure
URLDownloadErrorDescription = "ErrDownloadFailure"
Case ErrAuthenticationRequired
URLDownloadErrorDescription = "ErrAuthenticationRequired"
Case ErrNoValidMedia
URLDownloadErrorDescription = "ErrNoValidMedia"
Case ErrConnectionTimeout
URLDownloadErrorDescription = "ErrConnectionTimeout"
Case ErrInvalidRequest
URLDownloadErrorDescription = "ErrInvalidRequest"
Case ErrUnknownProtocol
URLDownloadErrorDescription = "ErrUnknownProtocol"
Case ErrSecurityProblem
URLDownloadErrorDescription = "ErrSecurityProblem"
Case ErrCannotLoadData
URLDownloadErrorDescription = "ErrCannotLoadData"
Case ErrCannotInstantiateObject
URLDownloadErrorDescription = "ErrCannotInstantiateObject"
Case ErrRedirectFailed
URLDownloadErrorDescription = "ErrRedirectFailed"
Case ErrRedirectToDir
URLDownloadErrorDescription = "ErrRedirectToDir"
Case ErrCannotLockRequest
URLDownloadErrorDescription = "ErrCannotLockRequest"
Case Else
URLDownloadErrorDescription = "ErrErrorUnknown"
End Select
GetDownloadErrorDescription = URLDownloadErrorDescription
End Function

'///////////////////////////////////////////////////////////////


Public Function BrowseForFolder(folder As String)
Dim obj, objFld, objPath
Dim objOrigen

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


On Error GoTo noSel
Set obj = CreateObject("Shell.Application")
Set objOrigen = obj.NameSpace("C:\")
Set objFld = obj.BrowseForFolder(Form1.hwnd, "Buscando Carpeta :", BIF_NEWDIALOGSTYLE Or BIF_RETURNONLYFSDIRS, folder)
Set objPath = objFld.Items.item

BrowseForFolder = objPath.Path
Exit Function

noSel:
BrowseForFolder = ""
End Function

Public Sub OnTop()
Dim ok As Long
'ok = SetWindowPos(Form1.hwnd, HWND_TOPMOST, Form1.ScaleLeft, Form1.ScaleTop, meWidth, meHeight, 0)
Debug.Print ok
End Sub

'///////////////////////////////////////////////////////////////
Function getTitle(file As String)
Dim sTitulo As String
Dim regexp As String
Dim texto As String
texto = leerKb(file)
sTitulo = getTitlefromText(texto)

sTitulo = Replace(sTitulo, ":", "-")
sTitulo = Replace(sTitulo, "?", "!")
sTitulo = Replace(sTitulo, """", "'")
sTitulo = Replace(sTitulo, "/", "-")
sTitulo = Replace(sTitulo, "\", "-")
sTitulo = Replace(sTitulo, "|", "-")
sTitulo = Replace(sTitulo, "<", "[")
sTitulo = Replace(sTitulo, ">", "]")
sTitulo = Replace(sTitulo, "*", "+")

getTitle = sTitulo
End Function

Function leerKb(file As String)
Dim obj As Object, txtstream As Object
Dim cadena As String
Set obj = CreateObject("Scripting.FileSystemObject")
Set txtstream = obj.OpenTextFile(file)
cadena = txtstream.read(1024)
leerKb = cadena
End Function

Function getTitlefromText(cadena As String)
Dim titulo As String
Dim pos1 As Integer, pos2 As Integer
'regexp = "<title>\s*.+\s*</title>"
pos1 = InStr(1, cadena, "<title>", vbTextCompare)
pos2 = InStr(pos1, cadena, "</title>", vbTextCompare)
pos1 = pos1 + Len("<title>")
titulo = Mid(cadena, pos1, pos2 - pos1)
titulo = Replace(titulo, vbCrLf, "")
getTitlefromText = titulo
End Function

'///////////////////////////////////////////////////////////////
Function esWeb(fichero As String) As Boolean
Dim ext As String, pos As Integer
esWeb = False
ext = Right(fichero, InStrRev(fichero, ".", , vbBinaryCompare))
MsgBox "EXTENSION:" & vbCrLf & ext
pos = InStrRev(ext, "asp", , vbTextCompare)
If pos > 0 Then esWeb = True
pos = InStrRev(ext, "htm", , vbTextCompare)
If pos > 0 Then esWeb = True
pos = InStrRev(ext, "php", , vbTextCompare)
If pos > 0 Then esWeb = True
End Function


Function getFreeNumber(file As String)
Dim n As Integer, newname As String, ext As String
getFreeNumber = file
n = 1
Do While Dir$(getFreeNumber, vbNormal) <> ""
    getFreeNumber = onlyName(file) & "(" & n & ")." & fileext(file)

Loop

End Function

Function onlyName(file As String)
Dim pos As Integer
pos = InStrRev(file, ".")
onlyName = Left(file, pos - 1)
End Function

Function fileext(file As String)
Dim pos As Integer
pos = InStrRev(file, ".")
fileext = Right(file, Len(file) - pos)
End Function

Public Function baseName(URL As String) 'only for urls//
baseName = Right$(URL, Len(URL) - InStrRev(URL, "/"))
End Function


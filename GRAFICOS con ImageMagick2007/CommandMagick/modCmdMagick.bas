Attribute VB_Name = "modCmdMagick"
Const convert = "convert.exe"
Const identify = "identify.exe"
Const debugTemp = "c:\temp\debugVerbose.txt"
Const vbCom As String = """"

Public myFile As String
Public myBackupDir As String
Public ImagesEfectos(20) As String
Public Const maxEfectos = 20
Public numEfectoActual As Integer

Public Sub Main()
getTempDir
scomand$ = Command$
If scomand$ = "" Then
    Form1.Visible = True
    Exit Sub
Else
    Form1.Visible = True
    Form1.gcm1.external = True
    Form1.gcm1.LoadNew sinComillas(scomand)
End If
End Sub


Function sinComillas(file As String)
If Left(file, 1) = vbCom Then
pos1 = InStr(1, file, vbCom)
pos2 = InStr(pos1 + 1, file, vbCom)
sinComillas = Mid(file, pos1 + 1, pos2 - pos1 - 1)
Else
    sinComillas = file
End If
End Function

Function basename(file As String)
basename = Right(file, Len(file) - InStrRev(file, "\"))
End Function

Function basepath(file)
pos = InStrRev(file, "\")
basepath = Left(file, pos)
End Function

Function executeMagick(comando As String, original As String, destino As String, Optional visibility As Integer) As Boolean
Dim doMonitor As String, dopause As String
Dim WshShell As Object
Screen.MousePointer = vbHourglass
doMonitor = " -monitor "
If visibility > 1 Then
    dopause = " & echo. & echo Error: %errorlevel% & pause"
End If
If Dir$(original, vbNormal) = "" Then GoTo NOORIGEN
executecmd = "cmd /C " & convert & doMonitor & " """ & original & """ " & comando & " """ & destino & """" & dopause
Set WshShell = CreateObject("WScript.Shell")
r = WshShell.Run(executecmd, 5, True)
If Dir$(destino, vbNormal) = "" Then GoTo NODESTINO
executeMagick = True
Screen.MousePointer = vbArrow
Exit Function
NOORIGEN:
Screen.MousePointer = vbArrow
MsgBox "Fichero ORIGEN No Existe" & vbCrLf & original, vbCritical, "EXECUTE MAGICK"
executeMagick = False
Exit Function
NODESTINO:
Screen.MousePointer = vbArrow
MsgBox "No se produjo la SALIDA. " & vbCrLf & "Salida Esperada:" & vbCrLf & destino, vbCritical, "EXECUTE MAGICK"
executeMagick = False
Exit Function

End Function

Public Sub doFileMini(pixels, original As String, destino As String)
mycomando$ = "-thumbnail """ & pixels & "x" & pixels & ">"""
'myMiniOriginal$ = myBackupDir & "\" & filename
executeMagick mycomando, original, destino, 0
End Sub

Function BackupIt(orig, bakup)
FileCopy orig, bakup
End Function


Function getTempDir()
Set filesys = CreateObject("Scripting.FileSystemObject")
Set tempfolder = filesys.GetSpecialFolder(2) ' temp folder
'tempname = filesys.GetTempName
myBackupDir = tempfolder

End Function

Public Function getEfectFileName(origFile As String, numEfecto As Integer)
getEfectFileName = myBackupDir & "\" & basename(origFile$) & "_Ef(" & Format(numEfecto, "0##") & ").jpg"
End Function


Function duplyIt(file$)
On Error GoTo ErrCopy
duplyIt = getCopyName(file)
FileCopy file, duplyIt
Exit Function
ErrCopy:
MsgBox "No se pudo copiar la imagen. Tienes permisos de escritura?"
End Function

Function getCopyName(file$)
Dim num As Integer
num = 0
part1 = "Ef("
part2 = ") de "
bn = basename(file)
bd = basepath(file)
pos = InStr(1, bn, part1, vbBinaryCompare)
If pos > 0 Then
    pos2 = InStr(pos + 1, bn, part2, vbBinaryCompare)
    bn = Mid(bn, pos2 + Len(part2), Len(bn))
End If

xfile = bd & part1 & Format(num, "0##") & part2 & bn
Do While Dir$(xfile, vbNormal) <> ""
    num = num + 1
    xfile = bd & part1 & Format(num, "0##") & part2 & bn
Loop

getCopyName = xfile

End Function


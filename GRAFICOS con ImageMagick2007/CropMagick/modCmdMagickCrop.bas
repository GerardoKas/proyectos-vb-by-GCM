Attribute VB_Name = "modCmdMagick"
Const convert = "convert.exe"
Const identify = "identify.exe"
'comillas
Const vbCom As String = """"
'Archivo actual abierto
Public myFile As String
'el archivo de backup
Public myBackup As String
'si esta elegido hacer backup
'Public doBackup As Boolean
'si eligio copiar con otro nombre
Public myRecorte As String
Public doDebug As Boolean

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
pos = InStrRev(file, "\")
basename = Right(file, Len(file) - pos)
End Function

Function basepath(file)
pos = InStrRev(file, "\")
basepath = Left(file, pos)
End Function
Function mogrifyMagick(comando As String, original As String, Optional destino As String)
Dim doMonitor As String, dopause As String
Dim WshShell As Object
If destino = "" Then
    destino = original
End If
doMonitor = " -monitor "
If doDebug Then
dopause = " & echo. & echo Error: %errorlevel% & pause"
End If
executecmd = "cmd /C convert.exe " & doMonitor & " """ & original & """ " & comando & " """ & destino & """ " & dopause
Debug.Print executecmd
Set WshShell = CreateObject("WScript.Shell")
r = WshShell.Run(executecmd, 5, True)
'msgBox executecmd
End Function

Public Function recortar(X, Y, w, h)
mogrifyMagick "-crop " & w & "x" & h & "+" & X & "+" & Y, myFile
End Function

Public Function copiarRecorte(X, Y, w, h)
If myRecorte = "" Then
    MsgBox "No se eligio nombre del recorte"
    Exit Function
End If
mogrifyMagick "-crop " & w & "x" & h & "+" & X & "+" & Y, myFile, myRecorte

End Function

Function BackupIt(original, destino)
FileCopy original, destino
End Function

Public Sub restaura(nueva, backup)
FileCopy backup, nueva
Kill backup


End Sub
Function getBackupName(file)
Const bak = ".BAK"
Dim num As Integer
num = 0
xfile = file & "(" & num & ")" & bak
Do While Dir$(xfile, vbNormal) <> ""
    num = num + 1
    xfile = file & "(" & num & ")" & bak
Loop
getBackupName = xfile
End Function

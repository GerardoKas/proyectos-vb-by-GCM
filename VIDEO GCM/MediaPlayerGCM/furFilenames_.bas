Attribute VB_Name = "furFilenames_"
'posibilidad de clickear y ubicar en el disco duro con explorer.exe
'guardar varias listas
'autocargar y autosalvar lo ultimo abierto
'
'
'
'

Public Type Archivo
segundos As Long
'frames As Integer
Tamano As Long
'fecha As Date
nomArk As String
Directorio As String
unidad As String
llegada As Integer
duracion As String * 10
End Type

Public ListaArk2(10000) As Archivo
Public NumArks As Long
Public ACTUAL As Integer
Const bar As String = "::"

'cadena tipica "C:\Archiv~1\Videos\imagen.avi"
'Solo se usa cuando se droppean los archivos en la lista
Public Function mkArkivo(cadena As String) As Archivo
Dim x As Integer
'x = InStr(1, cadena, "\", vbBinaryCompare)
mkArkivo.unidad = getUnidad(cadena)
'========0
ultimabarra = InStrRev(cadena, "\")
'========
mkArkivo.Directorio = getDirectorio(cadena) 'Mid(cadena, x, ultimabarra - x + 1)
mkArkivo.nomArk = Right(cadena, Len(cadena) - ultimabarra)
'If fileExists(cadena) Then mkArkivo.Tamano = FileLen(cadena)
End Function

Function getUnidad(text As String) As String
If Left(text, 2) = "\\" Then
    getUnidad = Left(text, InStr(3, text, "\") - 1)
ElseIf Mid(text, 2, 1) = ":" Then
    getUnidad = Left(text, 2)
End If
End Function

Function getDirectorio(text As String) As String
u = getUnidad(text)
getDirectorio = Replace(text, u, "", 1, 1)
getDirectorio = Left(getDirectorio, InStrRev(getDirectorio, "\"))
End Function


Public Function getFilename(elark As Archivo) As String
getFilename = elark.unidad & elark.Directorio & elark.nomArk
End Function

Public Function SecsToTime(segundos As Long) As String
Dim mins As Double, hrs As Integer, secs As Single
mins = segundos \ 60
secs = Int(segundos - mins * 60)
hrs = mins / 60
mins = mins - hrs * 60
SecsToTime = Format(hrs, "0#") & ":" & Format(mins, "0#") & ":" & Format(secs, "0#")
End Function

Public Function TheSize(size As Long) As String
If size / 1024 > 1024 Then
TheSize = Format((size / 1024) / 1024, "###.#") & "Mb."
ElseIf size < 1024 Then
TheSize = size & "Bytes"
Else
TheSize = Format(size / 1024, "###.#") & "Kb."
End If
End Function

Sub Savelist(file As String)
f = FreeFile()
Open file For Output As #f
For i = 0 To NumArks - 1
If ListaArk2(i).Directorio = "" Then Exit For
cad = ListaArk2(i).unidad & "" & ListaArk2(i).Directorio & ListaArk2(i).nomArk & bar & ListaArk2(i).segundos & bar & ListaArk2(i).Tamano
Print #f, cad
Next
Close #f
End Sub

Function loadlist(file As String) As Boolean
Dim cad As String, fl As String
If fileExists(file) = False Then
loadlist = False
Exit Function
End If
loadlist = False
f = FreeFile
Open file For Input As #f
i = 0
While Not EOF(f)
Line Input #f, cad
'p = InStr(1, cad, bar)
'fl = Left(cad, p - 1)
'p2 = InStr(p + Len(bar), cad, bar)
'secs = Right(cad, Len(cad) - p - 1)
pts = Split(cad, bar)
ListaArk2(i) = mkArkivo(CStr(pts(0)))
ListaArk2(i).segundos = Val(pts(1))
ListaArk2(i).Tamano = pts(2)
i = i + 1
Wend
Close #f
NumArks = i
loadlist = True
End Function

Public Function fileExists(file As String) As Boolean
On Error GoTo erro
If Dir$(file, vbNormal) <> "" Then
    fileExists = True
Else
    fileExists = False
End If
Exit Function
erro:
fileExists = False
End Function

Function checkPath(raiz As String) As Boolean
Dim sh As New Shell32.Shell, fol As Shell32.Folder
On Error GoTo noe
checkPath = False
'Set fol = sh.NameSpace(raiz)
'If fol Is Nothing Then
'    checkPath = False
'Else
'    checkPath = True
'End If
'''''''''''''''''
sh.Open (raiz)
If Dir$(raiz, vbNormal) <> "" Then
checkPath = True
End If
Exit Function
noe:
checkPath = False
End Function

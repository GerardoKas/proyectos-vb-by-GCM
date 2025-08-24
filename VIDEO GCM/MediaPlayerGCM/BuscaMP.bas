Attribute VB_Name = "Busca"
Option Explicit

Rem ESTA ES LA TAN REPETITIVA MODULA

Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type

' Const MAX_PATH = 260
 Const MAXDWORD = &HFFFF  'Para calcular el tamaño -"""High*DWord+Low"""-
 Const INVALID_HANDLE_VALUE = -1
 Const FILE_ATTRIBUTE_ARCHIVE = &H20  'comprimido
 Const FILE_ATTRIBUTE_DIRECTORY = &H10
 Const FILE_ATTRIBUTE_HIDDEN = &H2
 Const FILE_ATTRIBUTE_NORMAL = &H80
 Const FILE_ATTRIBUTE_READONLY = &H1
 Const FILE_ATTRIBUTE_SYSTEM = &H4
 Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public tmpLeyendo As Label


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function nonuls(cadena As String) As String
nonuls = Left(cadena, InStr(1, cadena, Chr(0), vbBinaryCompare) - 1)
End Function


'Esta cada vez se complica maaaaaaaaaaaaaaaaassssssshhhhh
'porque hay que guardar varios datos, el id de carpeta, la carpeta anterior etc

'La unidad en que se busca, el directorio, el pattron solo extension o estilo msdos (*da igual _
pq es un instr(1,cadenadeExtensiones,Extension)
'recomendable que sean extensiones separadas por espacios sin *.


Public Function BUSKA3(unidad As String, raiz As String, recursivo As Boolean, Buska As String, Archivos() As Archivo, cantidad As Long, sudir As Integer) As String
Attribute BUSKA3.VB_UserMemId = 0
Dim Fdl As String 'el nombre de archivo
Dim H As Long 'el id de la busqueda
Dim wf As WIN32_FIND_DATA 'los valores del archivo
Dim s As Long 'wf.cfilename de find-file
Dim Ext As String 'la extension
'Dim Anterior As Integer
'Static Tope As Integer
'Static Idcarp As Integer
'Static sudir As Integer

If Right(unidad, 1) = "\" Then unidad = Left(unidad, Len(unidad) - 1)
If Right(raiz, 1) <> "\" Then raiz = raiz + "\"
If Left(raiz, 1) <> "\" Then raiz = "\" + raiz

'Anterior = Idcarp
'Idcarp = Tope
DoEvents

H = FindFirstFile(unidad & raiz & "*" & Chr(0), wf)
If H <> -1 Then s = 1
tmpLeyendo = raiz
Do While s <> 0
    DoEvents
    If Left(wf.cFileName, 1) <> "." Then
    Fdl = nonuls(wf.cFileName)
    
    If wf.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
        'Debug.Print "DIR:" & Fdl
        If Fdl <> "." And Fdl <> ".." And Fdl <> "RECYCLED" Then
            'llamada recursiva
            sudir = sudir + 1
            'Debug.Print "DIRECTORIO : " & Fdl
            If recursivo Then Call BUSKA3(unidad, raiz + Fdl, recursivo, Buska, Archivos(), cantidad, sudir)
            'Idcarp = Idcarp - 1
            sudir = sudir - 1
        End If
    Else
        Ext = Exten(Fdl)
        'el problema ahora es guardar los archivos con su carpeta y su subdirectorio y demas.........
        If InStr(1, LCase(Buska), LCase(Ext)) And Len(Ext) >= 3 Then
              
              Archivos(cantidad).Directorio = raiz
              Archivos(cantidad).nomArk = Fdl
              Archivos(cantidad).unidad = unidad
              'Al añadir el tamaño al listview ya se convierte a kb o mb
              Archivos(cantidad).Tamano = FileLen(unidad & raiz & Fdl)
              

            cantidad = cantidad + 1
        End If
     'If Archivos(cantidad) = "" Then MsgBox "ohoh!muestra dos que no son ??"
   End If
   End If
s = FindNextFile(H, wf)
Loop
s = FindClose(H)
'mostrar.Caption = ""
End Function


Private Function Exten(texto As String) As String
Dim fin As String, pos As Integer
'fin = Right(texto, 4)
'pos = InStr(fin, ".")
'Exten = Mid(fin, pos + 1, Len(fin) - pos)
Exten = Right(texto, Len(texto) - InStrRev(texto, "."))
End Function

Public Function LaUNIDAD(path As String) As String
Dim x As Integer
x = InStr(1, path, "\")
LaUNIDAD = Left(path, x)
End Function


Attribute VB_Name = "rutinasForm"
Public myFolder As String

Private wsh As Object
Private fso As Object
Private sh As New Shell32.Shell
Private isFS As Boolean
Public HALTIT As Boolean
Public calculateFolderSize As Boolean

Dim carpetas(20) As String
Dim numcarpetas As Integer

Public Const myApp = "GCM_Explorer"
Public Const mySection = "Carpetas"
Public Const myConfig = "GCM_Explorer.ini"

Public Enum columnas
     eNombre = 1
     eTamano = 2
     eCreado = 3
     eTipo = 4
     eLength = 5
End Enum


Sub addColumns(lv As listview)
lv.ColumnHeaders.Add 1, "Nombre", "Nombre"
lv.ColumnHeaders.Add 2, "ORDER-Tamano", "Tamaño"
lv.ColumnHeaders.Add 3, "Creado", "Creado"
lv.ColumnHeaders.Add 4, "Tipo", "Tipo"
lv.ColumnHeaders.Add 5, "Length", "NameLength"
lv.ColumnHeaders.Add 6, "NUMBER-Tamano", "BYTES"
lv.ColumnHeaders.Add 7, "REAL-Path", "PATH"
lv.ColumnHeaders.Add 8, "REAL-TYPE", "TYPE"
lv.ColumnHeaders(5).Width = 0
lv.ColumnHeaders(6).Width = 0
'lv.ColumnHeaders(7).Width = 0
End Sub

Public Sub loadFolders(combo As ComboBox)
'If leerConfig(combo) = False Then
crearDefecto combo
'combo.AddItem "C:\"
'combo.AddItem "D:\"'
'combo.AddItem "E:\"
'End If
'Abrir la primera carpeta
combo.ListIndex = 0
End Sub

Public Function leerConfig(combo As ComboBox) As Boolean
f = FreeFile
If Dir$(myConfig) = "" Then leerConfig = False: Exit Function
Open myConfig For Input As #f
Do While Not EOF(f)
Line Input #f, linea
On Error Resume Next
If IsError(Dir$(linea)) = True Then
combo.AddItem linea
ElseIf Dir$(linea) <> "" Then
combo.AddItem linea
End If
Loop
Close #f
leerConfig = True
End Function

Public Function saveConfig(combo As ComboBox) As Boolean
f = FreeFile
Open myConfig For Output As #f
For i = 0 To combo.ListCount - 1
    Print #f, combo.List(i)
Next
Close #f
End Function

Public Sub crearDefecto(combo As ComboBox)
Dim wsh As Object
Dim objFolders As Object
Dim fso As Object
Dim pt As Object
Set wsh = CreateObject("WScript.Shell")
Set objFolders = wsh.specialfolders
Set fso = CreateObject("Scripting.FileSystemObject")
'Añadir todos al combo
For Each i In objFolders
combo.AddItem i
Next
Dim sh As New Shell32.Shell
tt = sh.NameSpace(ssfPERSONAL).Self.path
combo.AddItem tt
numcarpetas = objFolders.Count
End Sub


'Sub readDirOriginal(path As String, lv As listview)
'Dim itemX As ListItem
'lv.ListItems.Clear
'If Right(path, 1) <> "\" Then path = path & "\"
'Set fso = CreateObject("Scripting.FileSystemObject")
'On Error GoTo noExiste
'Set folder = fso.GetFolder(path)
'myFolder = path
'On Error Resume Next
'Set subfolders = folder.subfolders
'
'lv.View = lvwReport
'For Each file In subfolders
'    If Not (file = "." Or file = ".." Or file = "") Then
'    Set itemX = lv.ListItems.Add(, "@" & file.Name, file.Name)
'    itemX.SmallIcon = loadFile(path & file.Name)
'    'itemX.SubItems(2) = Format(file.Size, "###,###.00")
'    itemX.SubItems(2) = fecha(file.datecreated)
'    itemX.SubItems(3) = file.Type
'    itemX.SubItems(4) = Len(file.Name)
'    End If
'Next
'
'Set Files = folder.Files
'For Each file In Files
'    Set itemX = lv.ListItems.Add(, "@" & file.Name, file.Name)
'    If itemX Is Nothing Then MsgBox "no se hixo " & file.Name
'    itemX.SmallIcon = loadFile(path & file.Name)
'    itemX.SubItems(1) = Tamano(file.size)
'    itemX.SubItems(2) = fecha(file.datecreated)
'    itemX.SubItems(3) = file.Type
'    itemX.SubItems(4) = Len(file.Name)
'
'Next
'
'Exit Sub
'noExiste:
'MsgBox "No se ha encontrado la ruta : " & vbCrLf & path, vbExclamation + vbOKOnly, "ERROR"
'End Sub

Function fecha(d As Date) As String
fecha = Format(d, "dd/mm/yy", vbMonday, vbFirstJan1)
End Function

Function Tamano(size As Long) As String
Dim pc As Double
Dim tm As String
If size < 1024 Then

Tamano = size & " Bytes"
ElseIf size < 1048576 Then

Tamano = fmtSingle(size / 1024) & " Kb"
ElseIf size < 1073741824 Then
Tamano = fmtSingle(size / 1024 / 1024) & " Mb"
ElseIf size < 1099511627776# Then
Tamano = fmtSingle(size / 1024 / 1024 / 1024) & " Gb"
End If

End Function

Function fmtSingle(dato As Single)
fmtSingle = Format(dato, "####.##")
End Function
Function bitTamano(size As Long) As String
bitTamano = Format(size, "0#########")
End Function

Sub readDir(path As String, lv As listview)
'On Local Error GoTo Exo
Dim sh As New Shell
Dim itemX As ListItem
Dim tmn As String
Dim folder As folder
Dim file As FolderItem
Dim fullName As String
If Not esFolder(path) Then Exit Sub
If isFS = True Then
If Right(path, 1) <> "\" Then path = path & "\"
End If
'calculateFolderSize = True
lv.ListItems.Clear
lv.Visible = False
DoEvents

myFolder = path
HALTIT = False


'por vbscript
If calculateFolderSize = True Then
Set fso = CreateObject("Scripting.FileSystemObject")
'Set folder = fso.GetFolder(path)
End If
Form1.lblDesc.Caption = "Cargando " & path
Set folder = sh.NameSpace(path)
If folder Is Nothing Then MsgBox "nn fld": Exit Sub
'f vts fs bs
Exit Sub

Set subfolders = folder.Items 'folder.subfolders
On Error GoTo DONE
Form1.lblDesc.Caption = "Leyendo " & folder.Title
For Each file In subfolders
    If HALTIT = True Then
    'terminar abruptamente
        Set itemX = lv.ListItems.Add(1, "STOPPED", "HALTED!", 0, 0)
        itemX.SubItems(1) = "DIDNT FINISH!!"
        itemX.ListSubItems(1).Bold = True
        GoTo DONE
    End If
    
    DoEvents
    
    Form1.lblDesc.Caption = ">" & file.Name
    
    'realName = folder.Self.path
    mtype = ""
    Debug.Print "FILE::" & file.Name
    Debug.Print file.path
    If file.IsFolder Then mtype = "FLD ": Debug.Print "is folder"
    If file.IsBrowsable Then mtype = mtype & "BROWS ": Debug.Print "browsable"
    If file.IsFileSystem Then mtype = mtype & "FS ": Debug.Print "filesystem"
    If file.IsLink Then mtype = mtype & "LNK ": Debug.Print "is link"
    Debug.Print file.Type
    Debug.Print vbCrLf
    
    'si es carpeta pero no es un zip
    If file.IsFolder And Not file.IsBrowsable Then
'folder tb son los zip
'isfilesystem solo del discoduro

        Set itemX = lv.ListItems.Add(, "@" & file.Name, file.Name)
        itemX.SmallIcon = loadFile(file.path)
        If calculateFolderSize = True Then
        'debemos usar vbsscript
            Form1.lblDesc.Caption = "Calculando Tamano:" & file.Name
            Set fldvbs = fso.GetFolder(file.path)
            DoEvents
            itemX.SubItems(1) = Tamano(fldvbs.size)
            itemX.SubItems(5) = bitTamano(fldvbs.size)
        End If
        If file.IsBrowsable Then
        MsgBox "como es que entra aca si no es browsable"
            itemX.SubItems(1) = Tamano(file.size)
            itemX.SubItems(5) = bitTamano(file.size)
        End If
        itemX.SubItems(2) = fecha(file.ModifyDate)
        itemX.SubItems(3) = file.Type
        
        'itemX.ListSubItems(3).Tag =
        
        itemX.SubItems(4) = Format(Len(file.Name), "0##")
    Else
        DoEvents
        'añladir columna de real name (clsid,lnk...) vs. friendly name
        fn = file.Name
        If file.IsLink Then
            fn = file.Name & ".lnk"
            Tag = "Link"
'            realName = file.Name & ".lnk"
'        Else
'            fn = file.Name
'            Tag = fn
        End If
     Set itemX = lv.ListItems.Add(, "@" & fn, fn)
     '   itemX.ToolTipText = Tag
        'If itemX Is Nothing Then MsgBox "no se hixo " & file.Name
        itemX.SmallIcon = loadFile(file.path)
        itemX.SubItems(1) = Tamano(file.size)
        itemX.SubItems(2) = fecha(file.ModifyDate)
        itemX.SubItems(3) = file.Type
        itemX.SubItems(4) = Format(Len(file.Name), "0##")
        itemX.SubItems(5) = bitTamano(file.size)
    End If
    itemX.SubItems(6) = file.path
'        If file.IsBrowsable Then
'        itemX.ListSubItems(3).ForeColor = RGB(255, 200, 200)
'        End If
'        If file.IsFolder Then
'            'itemX.ListSubItems(3).ForeColor = RGB(200, 255, 200)
'        End If
'        If file.IsFileSystem Then
'           ' itemX.ListSubItems(3).ForeColor = RGB(255, 255, 200)
'        End If
'        If file.IsLink Then
'             itemX.Bold = True
'        End If
        itemX.SubItems(7) = mtype
Next
DONE:
lv.Visible = True
DoEvents
Exit Sub
Exo:
MsgBox "Error!" & vbCrLf & Err.Number & "::" & Err.Description & vbCrLf & "ORIGEN:" & Err.Source & " "
End Sub

Sub nombresLargos(lv As listview, longitud As Integer)
Dim item As ListItem
For Each item In lv.ListItems
    If Len(item.Text) > longitud Then
        item.Bold = True
    End If
Next
End Sub

'Function getVBSObject(filename As String) As Object
'Dim fso As Object
'Set fso = CreateObject("Scripting.FileSystemObject")
'If esFolder(filename) Then
'    Set getVBSObject = fso.GetFolder(filename)
'Else
'    Set getVBSObject = fso.getFile(filename)
'End If
'End Function

'Function getShellObject(filename As String) As Object
'Dim shl As New Shell
'If (esFolder(filename)) Then
'    Set getShellObject = shl.NameSpace(filename)
'Else
'    Dim folder As folder
'    Set folder = shl.NameSpace(basepath(filename))
'    Set getShellObject = folder.ParseName(basename(filename))
'End If
'End Function
Function esFolder(path As String) As Boolean
Dim shl As New Shell, fld As folder
On Error Resume Next
isFS = True
Set fld = shl.NameSpace(path)
If fld Is Nothing Then esFolder = False: Exit Function
isFS = fld.Self.IsFileSystem
If fld.Self.IsFolder = True Then esFolder = True: Exit Function
esFolder = False
End Function

'Function esCLSID(cadena As String)
'Dim shl As New Shell, fld As folder, item As FolderItem
'On Error GoTo Fin
'Set fld = shl.NameSpace(cadena)
'If fld.Self.IsBrowsable = True And fld.Self.IsFolder = False Then
'    esCLSID = True
'    Exit Function
'End If
'Fin:
'esCLSID = False
'End Function

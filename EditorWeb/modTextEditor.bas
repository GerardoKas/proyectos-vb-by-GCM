Attribute VB_Name = "modTextEditor"
Public fso As FileSystemObject
Public fsoFile As TextStream

Public OriginalFile As String
Public tempFile As String

Sub showTextFile(file As String)
If fso.FileExists(file) = False Then
    MsgBox "No existe " & file
End If
Form1.txtContents.Text = getTextFile(file)
End Sub

Sub insertar(texto)
parte1 = Left(txtContents.Text, txtContents.SelStart)
parte2 = Right(txtContents.Text, Len(txtContents.Text) - txtContents.SelStart)
txtContents.Text = parte1 & texto & parte2
End Sub

Function getTextFile(file As String) As String
On Error Resume Next
Set openfile = fso.OpenTextFile(file)
getTextFile = openfile.ReadAll()
End Function

Function getTempName(Optional fichero)
Dim mDir As String, mfile As String
If IsMissing(fichero) Then  ' pide un temporal completo
    'mDir = fso.GetSpecialFolder(TemporaryFolder)
    mDir = App.path
    mfile = fso.getTempName & ".htm"
Else
    mDir = fso.GetParentFolderName(fichero)
    mfile = fso.GetBaseName(fichero) & ".TEMP.htm"
End If
If testWrite(mDir) = False Then
    mDir = fso.GetSpecialFolder(TemporaryFolder)
End If
getTempName = IIf(Right(mDir, 1) = "\", mDir, mDir & "\") & mfile
tempFile = getTempName
End Function

Sub saveTextFile(filename, ByRef texto As String)
If filename = "" Then
MsgBox "Mea culpa. No se creo el nombre de fichero. VUELVE A INTENTARLO"
Exit Sub
End If
Set writefile = fso.OpenTextFile(filename, ForWriting, True)
writefile.Write texto
writefile.Close
End Sub

Public Sub Duplicados(combo As ComboBox)
Dim Search1 As Long
Dim Search2 As Long
Dim KillDupe As Long
    KillDupe = 0
    For Search1& = 0 To combo.ListCount - 1
    For Search2& = Search1& + 1 To combo.ListCount - 1
    KillDupe = KillDupe + 1
    If combo.List(Search1&) = combo.List(Search2&) Then
        combo.RemoveItem Search2&
        Search2& = Search2& - 1
    End If
    Next Search2&
    Next Search1&
End Sub

Sub makeBackup(original, backup)
Debug.Print "ORIGINAL:" & original
Debug.Print "BACKUP:" & backup
On Error GoTo ero
FileCopy original, backup
Exit Sub
ero:
MsgBox "No se creo backup.." & vbCrLf & " Tal vez pq es nuevo...", vbInformation + vbOKOnly, "NOSE"

End Sub

Function testWrite(path As String) As Boolean
On Error GoTo NOP
testfile = path & "\NONE.TMP"
f = FreeFile
Open testfile For Output As #f
Print #f, "TEST"
Close #f
Kill testfile
testWrite = True
Exit Function
NOP:
Close #f
testWrite = False
End Function

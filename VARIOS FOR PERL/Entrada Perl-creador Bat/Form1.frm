VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private datafile As String
Private ok As Long
Private Const myTitle As String = "Perl_Batcher"

Private Sub Form_Load()
Dim entrada As String
datafile = App.Path & "\" & "dataperl.dat"
entrada = Command$()
entrada = Replace(entrada, Chr(34), "")

If Dir$(entrada, vbArchive Or vbNormal) = "" Then
    MsgBox "Debes enviar un (UNICO) archivo para generar el batch de perl"
    ok = 0
Else
    If entrada = "" Then
        MsgBox "No hay entrada - Se Esperaba un archivo .pl"
        ok = 0
    Else
        crearbatch entrada
        ok = 1
    End If
End If
If ok = 0 Then
If (MsgBox("Deseas registrar el porgrama como menu para los archivos *.pl?", vbYesNo, "Registro") = vbYes) Then
    registrar
End If
End If
Unload Me
End

End Sub

Sub crearbatch(myFile As String)
Dim Data As String, textobat As String
Data = abrirDat()
If Data <> "" Then
    textobat = Replace(Data, "%ARCHIVO%", myFile)
    grabar textobat, myFile
Else
    ok = 0
End If

End Sub

Function abrirDat()
Dim kb As Long
Dim Data As String
kb = FileLen(datafile)
file = FreeFile()
If (Dir$(datafile) <> "") Then
    Open datafile For Input As #file
    Data = Input(kb, #file)
    Close #file
    abrirDat = Data
Else
    MsgBox "Debe existir el fichero " & datafile & " con la estructura del batch y %ARCHIVO%"
    abrirDat = 0

End If
End Function

Sub grabar(texto As String, filename As String)
ffile = FreeFile()

Open filename & ".bat" For Output As #ffile
Print #ffile, texto
Close #ffil
End Sub

Sub registrar()
Dim valor As String, ext As String, ok As String, clave As String
valor = App.Path & "\" & App.EXEName & " ""%1""" & Chr(0)
myKey = HKEY_CLASSES_ROOT
'SetDefaultValue myKey, "Shell\Batcher\command", valor

ext = RegGetstring(HKEY_CLASSES_ROOT, ".pl", "")
If Not itsok(ext) Then
    MsgBox "No se encontro asociacion para .pl"
    Exit Sub
End If
clave = ext & "\Shell\Batcher\command"
ok = RegSavestring(HKEY_CLASSES_ROOT, clave, "", valor)

If itsok(ok) Then
    MsgBox "Se ha registrado la Aplicacion para archivos pl" & vbCrLf & "Bajo la extension " & ext & " en la ruta " & clave
Else
    MsgBox "Ha ocurrido un error al registrar el programa con " & clave
End If

End Sub


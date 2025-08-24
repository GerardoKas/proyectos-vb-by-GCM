Attribute VB_Name = "modWebEditor"
'Public myDoc As HTMLDocument
Public myelement As IHTMLElement
Public myIniTags As String
Public myDoc As HTMLDocument
Public regEx As Object
Public basepath As String
Public goingOk As Boolean

Sub showWebFile(fichero)
Form1.web1.Stop
If fso.FileExists(fichero) = False Then
    'Form1.web1.Navigate "about:blank"
    MsgBox "NO EXISTE " & fichero
    'Exit Sub
Else
GoToDir fichero
'Form1.web1.SetFocus
End If
DoEvents
End Sub

Sub GoToDir(direccion)
goingOk = True
Form1.web1.Silent = True
Form1.web1.Navigate direccion
End Sub

Function getWebHtml()
'On Error Resume Next
getWebHtml = myDoc.body.parentElement.outerHTML
End Function

Function myEvent() As IHTMLEventObj
Set myEvent = myDoc.parentWindow.event
End Function

Function loadTags(combo As ComboBox, iniFile As String)
Dim stag As String
myIniTags = iniFile
ReadWrite.setIniFile myIniTags
stag = ReadWrite.ReadKeys("TAGS")
tags = Split(stag, Chr(0))
combo.AddItem "Etiqueta"
For Each i In tags
    If i <> "" Then
    combo.AddItem i
    End If
Next
combo.ListIndex = 0
End Function

Function getTagCode(tag As String)
ReadWrite.setIniFile myIniTags
scode = ReadWrite.ReadFromFile("TAGS", tag)
getTagCode = scode
End Function

Function getRelativePath(webPath, pagePath)
'f = "file:///"
'redPath = Replace(webPath, f, "")
getRelativePath = Replace(webPath, pagePath, "")
End Function

Function replacePaths(texto) As String
Dim vbreg As RegExp
Dim matches As MatchCollection, match As match
Set vbreg = New RegExp
vbreg.Pattern = "(""|')file:///([^""']+)(""|')"
vbreg.MultiLine = True
vbreg.Global = True
Set matches = vbreg.Execute(texto)
'MsgBox matches.Count
For Each i In matches
    Debug.Print i
    texto = Replace(texto, i, getRelativePath(i, myDoc.URL))
Next
replacePaths = texto
End Function

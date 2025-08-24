Attribute VB_Name = "furFilenames_"
Public Type Archivo
nomArk As String
Directorio As String
Unidad As String
End Type
'cadena tipica "C:\Archiv~1\Videos\imagen.avi"

Public Sub rutadeark(cadena As String, ElArk As Archivo)
Dim x As Integer
x = InStr(1, cadena, ":\", vbBinaryCompare)
ElArk.Unidad = Left(cadena, x)

End Sub

Private Function ultimabarra(cadena As String) As Integer
For i = Len(cadena) To 1 Step -1
If (Mid(cadena, i, 1) = "\") Then
ultimabarra = i
Exit For
End If
Next
End Function

Public Function ELdirdelFile(cadena As String) As String
ELdirdelFile = Left(cadena, ultimabarra(cadena))

End Function

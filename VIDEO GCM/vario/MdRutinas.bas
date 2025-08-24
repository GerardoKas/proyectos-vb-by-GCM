Attribute VB_Name = "MdRUtinas"
Public Function Nsudir(ruta As String) As Integer
Dim num As Integer
Do
    pos = InStr(pos + 1, ruta, "\", vbBinaryCompare)
    num = num + 1
Loop While pos
If Right(ruta, 1) = "\" Then
Nsudir = num - 1
Else
Nsudir = num
End If

End Function

Public Function Updir(ruta As String) As String
Dim z As Integer, r As Integer
For z = Len(ruta) - 1 To 0 Step -1
If Mid(ruta, z, 1) = "\" Then r = r + 1

If r = 1 Then
    Updir = Left(ruta, z)
    Exit For
End If
Next
End Function

Public Function LaCarp(ruta As String, indice As Integer) As String
For i = 0 To indice
pos = InStr(pos + 1, ruta, "\", vbBinaryCompare)
Next

LaCarp = Left(ruta, pos)
End Function

Public Sub Ajustar(Imagen As Image, MaxAncho, MaxAlto)
Dim alto As Long, ancho As Long
Dim propor As Long, Anim As Long, Alim As Long

'Imagen.Visible = False
'Imagen.AutoSize = True
Imagen.Visible = False
Imagen.Stretch = False
'Imagen. = vbTwips 'vbPixels
Anim = Imagen.Width ' / 26.5
Alim = Imagen.Height ' / 26.5
If propor = 1 Then
    propor = Escalanow
ElseIf propor = 0 Then

    LadoImagenMayor = IIf(Imagen.Width > Imagen.Height, MaxAncho / Anim, MaxAlto / Alim)
End If

propor = LadoImagenMayor * 100
'Imagen.PaintPicture Imagen.Picture, 0, 20, Anim * (propor / 100), Alim * (propor / 100)
Imagen.Width = Anim * (propor / 100)
Imagen.Height = Alim * (propor / 100)
Imagen.Stretch = True
DoEvents
'Imagen.ScaleHeight = MaxAncho
'Imagen.ScaleWidth = MaxAncho


Imagen.Visible = True
End Sub

Public Function RutadArbol(nodepath As String) As String
Dim z As Integer, start As Integer, ruta As String
ruta = ""
start = 1
Do
z = InStr(start, nodepath, "\", vbBinaryCompare)
    If Len(Right(nodepath, Len(nodepath) - z)) > 0 Then
        RutadArbol = RutadArbol + Mid$(nodepath, start, z - start + 1)
        
    End If
    If z = Len(nodepath) Then
        RutadArbol = RutadArbol + Mid$(nodepath, start, z - start + 1)
    End If
        
    start = z + 2
Loop While start < Len(nodepath)

End Function

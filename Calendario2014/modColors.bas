Attribute VB_Name = "modColors"
Public colors As Variant
Public total As Integer

Public Function getColors()
If total = 0 Then
colors = Array("FFFFFF", "FFFF00", "FF0000", "00FF00", "00FFFF", "FF00FF", "0000FF", "000000")
total = UBound(colors)
End If

End Function

Public Function addColorHex(value As String)
For i = 0 To UBound(colors)
    If UCase(colors(i)) = UCase(value) Then Exit Function
Next
total = total + 1
ReDim Preserve colors(total)
colors(total) = value
End Function

Public Sub showColors(pic As PictureBox)
getColors
pic.ScaleMode = vbPixels
pic.DrawWidth = 16
pic.FillStyle = 0
For i = 0 To UBound(colors)
    c = colors(i)
    If c <> -1 Then
    pic.ForeColor = RGB(r(c), g(c), b(c))
    n = i * 16 + 8
    pic.Line (n, 0)-(n, pic.ScaleHeight)
    End If
Next
End Sub



Public Function toDec(value)
Dim valores
On Error GoTo ErrValue

valores = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, "A", "B", "C", "D", "E", "F")
If value = "" Then toDec = 0: Exit Function
v1 = Left(value, 1)
v2 = Right(value, 1)

For i = 0 To 16
    If UCase(v1) = CStr(valores(i)) Then
        vdec1 = i
        Exit For
    End If
Next

For i = 0 To 16
    If UCase(v2) = CStr(valores(i)) Then
        vdec2 = i
        Exit For
    End If
Next
toDec = vdec1 * 15 + vdec2
Exit Function
ErrValue:
MsgBox "Ha ocurrido un error al traducir el siguiente valor 'hexadecimal' :" & value & vbCrLf & "No se pueden leer nombres de color"
End Function

Public Function r(v)
r = Mid(v, 1, 2)
r = toDec(r)
End Function


Public Function g(v)
g = Mid(v, 3, 2)
g = toDec(g)
End Function


Public Function b(v)
b = Mid(v, 5, 2)
b = toDec(b)
End Function


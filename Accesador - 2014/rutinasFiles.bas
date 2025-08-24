Attribute VB_Name = "rutinasFiles"
Function basepath(filename As String) As String
Dim pos As Integer, desde As Integer
desde = -1
If Right(filename, 1) = "\" Then filename = Left(filename, Len(filename) - 1)
pos = InStrRev(filename, "\", -1, vbBinaryCompare)
basepath = Left(filename, pos)
End Function

Function basename(filename As String) As String
Dim pos As Integer
pos = InStrRev(filename, "\")
If pos = 0 Then basename = "": Exit Function
basename = Right(filename, Len(filename) - pos)
End Function

Function updir(path As String) As String
Dim pos As Integer
updir = basepath(path)
If updir = "" Then
updir = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
End If
End Function

Function formatKb(bytes As Long)

End Function

Function formatDate(fecha As Date)

End Function

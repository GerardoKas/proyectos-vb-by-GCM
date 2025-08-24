Attribute VB_Name = "Version2"
'   Visual Basic 6.0 - Use all of these API
'
'  Visual Basic 6.0   (Return To Forum Posts)    View Conversation    Convert C# to VB.NET    Archived Post
'Use all of these API
''''''''
'http://www.eggheadcafe.com/forums/ForumPost.asp?ID=50412&INTID=8

Private Declare Function lstrcpy Lib "kernel32" _
 Alias "lstrcpyA" _
 (ByVal lpString1 As String, _
 ByVal lpString2 As Long) As Long
 
Private Declare Sub MoveMemory Lib "kernel32" _
 Alias "RtlMoveMemory" _
 (dest As Any, _
 ByVal Source As Long, _
 ByVal Length As Long)
 
Private Declare Function VerQueryValue Lib "Version.dll" _
 Alias "VerQueryValueA" _
 (pBlock As Any, _
 ByVal lpSubBlock As String, _
 lplpBuffer As Any, _
 puLen As Long) As Long
 
Private Declare Function GetFileVersionInfo Lib "Version.dll" _
 Alias "GetFileVersionInfoA" _
 (ByVal lptstrFilename As String, _
 ByVal dwhandle As Long, _
 ByVal dwlen As Long, _
 lpData As Any) As Long
 
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" _
 Alias "GetFileVersionInfoSizeA" _
 (ByVal lptstrFilename As String, _
 lpdwHandle As Long) As Long
 
Function GetFileVersion(FullFileName As String) As String
 
Dim Buffer As String
Dim rc As Long
Dim lBufferLen As Long, lDummy As Long
Dim sBuffer()  As Byte
Dim lVerPointer As Long
Dim bytebuffer(255) As Byte
Dim Lang_Charset_String As String
Dim HexNumber As Long
Dim strVersionInfo(7) As String
Dim i As Integer
Dim strTemp As String
 
 '*** Get size ****
 lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
 If lBufferLen <= 1 Then
 Exit Function
 End If
 
 
 ReDim sBuffer(lBufferLen)
 rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
 If rc = 0 Then
 Exit Function
 End If
 
 rc = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", _
 lVerPointer, _
 lBufferLen)
 
 If rc = 0 Then
 Exit Function
 End If
 'lVerPointer is a pointer to four 4 bytes of Hex number,
 'first two bytes are language id, and last two bytes are code
 'page. However, Lang_Charset_String needs a  string of
 '4 hex digits, the first two characters correspond to the
 'language id and last two the last two character correspond
 'to the code page id.
 
 
 MoveMemory bytebuffer(0), lVerPointer, lBufferLen
 
 HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + _
 bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
 Lang_Charset_String = Hex(HexNumber)
 'now we change the order of the language id and code page
 'and convert it into a string representation.
 'For example, it may look like 040904E4
 'Or to pull it all apart:
 '04------        = SUBLANG_ENGLISH_USA
 '--09----        = LANG_ENGLISH
 ' ----04E4 = 1252 = Codepage for Windows:Multilingual
 
 Do While Len(Lang_Charset_String) < 8
 Lang_Charset_String = "0" & Lang_Charset_String
 Loop
 
 
 strVersionInfo(0) = "CompanyName"
 strVersionInfo(1) = "FileDescription"
 strVersionInfo(2) = "FileVersion"
 strVersionInfo(3) = "InternalName"
 strVersionInfo(4) = "LegalCopyright"
 strVersionInfo(5) = "OriginalFileName"
 strVersionInfo(6) = "ProductName"
 strVersionInfo(7) = "ProductVersion"
 
 For i = 0 To 7
 Buffer = String(255, 0)
 strTemp = "\StringFileInfo\" & Lang_Charset_String _
 & "\" & strVersionInfo(i)
 rc = VerQueryValue(sBuffer(0), strTemp, _
 lVerPointer, lBufferLen)
 
 If rc = 0 Then
 Exit Function
 End If
 
 lstrcpy Buffer, lVerPointer
 Buffer = Mid$(Buffer, 1, InStr(Buffer, Chr(0)) - 1)
 cad = cad & vbCrLf & Trim$(Buffer)
 Next i
 GetFileVersion = cad
End Function

Function Version2(file As String)
Version2 = GetFileVersion(file)
End Function

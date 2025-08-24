Attribute VB_Name = "Version1"
'**************************************
'Windows API/Global Declarations for :Ge
'     t Version Number for EXE, DLL or OCX fil
'     es
'**************************************
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Type FILEINFO
 CompanyName As String
 FileDescription As String
 FileVersion As String
 InternalName As String
 LegalCopyright As String
 OriginalFileName As String
 ProductName As String
 ProductVersion As String
 Description As String
 End Type
Public Enum VerisonReturnValue
 eOK = 1
 eNoVersion = 2
End Enum
'**************************************
' Name: Get Version Number for EXE, DLL
'     or OCX files
' Description:This function will retriev
'     e the version number, product name, orig
'     inal program name (like if you right cli
'     ck on the EXE file and select properties
'     , then select Version tab, it shows you
'     all that information) etc
' By: Serge
'
' Returns:FileInfo structure
'
' Assumes:Label (named Label1 and make i
'     t wide enough, also increase the height
'     of the label to have size of the form),
'     Common Dilaog Box (CommonDialog1) and a
'     Command Button (Command1)
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.1JavaStreet.com/vb/scripts/ShowCode.a
'     sp?txtCodeId=4976&lngWId=1'for details.'**************************************
Public Function GetFileVersionInformation(ByRef pstrFieName As String, ByRef tFileInfo As FILEINFO) As VerisonReturnValue
 Dim lBufferLen As Long, lDummy As Long
 Dim sBuffer() As Byte
 Dim lVerPointer As Long
 Dim lRet As Long
 Dim Lang_Charset_String As String
 Dim HexNumber As Long
 Dim i As Integer
 Dim strTemp As String
 'Clear the Buffer tFileInfo
 tFileInfo.CompanyName = ""
 tFileInfo.FileDescription = ""
 tFileInfo.FileVersion = ""
 tFileInfo.InternalName = ""
 tFileInfo.LegalCopyright = ""
 tFileInfo.OriginalFileName = ""
 tFileInfo.ProductName = ""
 tFileInfo.ProductVersion = ""
 '''
 lBufferLen = GetFileVersionInfoSize(pstrFieName, lDummy)
 If lBufferLen < 1 Then
 GetFileVersionInformation = eNoVersion
 Exit Function
 End If
 ReDim sBuffer(lBufferLen)
 lRet = GetFileVersionInfo(pstrFieName, 0&, lBufferLen, sBuffer(0))
 If lRet = 0 Then
 GetFileVersionInformation = eNoVersion
 Exit Function
 End If
 lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
 If lRet = 0 Then
 GetFileVersionInformation = eNoVersion
 Exit Function
 End If
 Dim bytebuffer(255) As Byte
 MoveMemory bytebuffer(0), lVerPointer, lBufferLen
 HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
 Lang_Charset_String = Hex(HexNumber)
 'Pull it all apart:
 '04------= SUBLANG_ENGLISH_USA
 '--09----= LANG_ENGLISH
 ' ----04E4 = 1252 = Codepage for Windows
 '     :Multilingual
 Do While Len(Lang_Charset_String) < 8
 Lang_Charset_String = "0" & Lang_Charset_String
 Loop
 Dim strVersionInfo(8) As String
 strVersionInfo(0) = "CompanyName"
 strVersionInfo(1) = "FileDescription"
 strVersionInfo(2) = "FileVersion"
 strVersionInfo(3) = "InternalName"
 strVersionInfo(4) = "LegalCopyright"
 strVersionInfo(5) = "OriginalFileName"
 strVersionInfo(6) = "ProductName"
 strVersionInfo(7) = "ProductVersion"
 strVersionInfo(8) = "Description"
 Dim Buffer As String
 For i = 0 To 7
 Buffer = String(255, 0)
 strTemp = "\StringFileInfo\" & Lang_Charset_String _
 & "\" & strVersionInfo(i)
 lRet = VerQueryValue(sBuffer(0), strTemp, _
 lVerPointer, lBufferLen)
 If lRet = 0 Then
 'GetFileVersionInformation = eNoVersion
 Buffer = "(NONE)" & Chr(0)
 'Exit Function
 Else
 lstrcpy Buffer, lVerPointer
 Buffer = Mid$(Buffer, 1, InStr(Buffer, vbNullChar) - 1)
 End If
 Select Case i
 Case 0
 tFileInfo.CompanyName = Buffer
 Case 1
 tFileInfo.FileDescription = Buffer
 Case 2
 tFileInfo.FileVersion = Buffer
 Case 3
 tFileInfo.InternalName = Buffer
 Case 4
 tFileInfo.LegalCopyright = Buffer
 Case 5
 tFileInfo.OriginalFileName = Buffer
 Case 6
 tFileInfo.ProductName = Buffer
 Case 7
 tFileInfo.ProductVersion = Buffer
 'Case 8
 'tFileInfo.Description = Buffer
 
 End Select
Next i
GetFileVersionInformation = eOK
End Function
'-----------
Function Version1(strFile As String)
 Dim udtFileInfo As FILEINFO
 On Error Resume Next

 If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
 datos = "No version available For this file"
 Else
 datos = "Company Name: " & udtFileInfo.CompanyName & vbCrLf
 datos = datos & "File Description:" & udtFileInfo.FileDescription & vbCrLf
 datos = datos & "File Version:" & udtFileInfo.FileVersion & vbCrLf
 datos = datos & "Internal Name: " & udtFileInfo.InternalName & vbCrLf
 datos = datos & "Legal Copyright: " & udtFileInfo.LegalCopyright & vbCrLf
 datos = datos & "Original FileName:" & udtFileInfo.OriginalFileName & vbCrLf
 datos = datos & "Product Name:" & udtFileInfo.ProductName & vbCrLf
 datos = datos & "Product Version: " & udtFileInfo.ProductVersion & vbCrLf
 End If
Version1 = datos
End Function

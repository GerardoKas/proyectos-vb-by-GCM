Attribute VB_Name = "ERWR"
'**************************************
' Name: EASIEST Read/Write to registry
' Description:This code makes it VERY ea
'     sy for you or a user to enter or recive
'     any part of the windows registry.
' By: Kevin Mackey
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None yet!
'
'**************************************


' Easiest Read/Write to Registry
' Kevin Mackey
' LimpiBizkit@aol.com
'Return Code add on's by James Blanchette
'itech@itechecom.com
'Begware Software ( a division of Independent Technical Services)
'http://www.itechecom.com
' We take no credit for any of the routines except for adding a return code to the functions
'to do an error check
' check the first 3 characters of the return string
'    if left(var,3)="NO! then
'      msgbox("error has occured")
'    End If


'
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long


Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long


Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    
    Public Const REG_SZ = 1 ' Unicode nul terminated string
    Public Const REG_DWORD = 4 ' 32-bit number




Public Function RegGetstring(hkey As Long, strPath As String, strValue As String)
    'EXAMPLE:
    '
    'text1.text = RegGetstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                RegGetstring = Left$(strBuf, intZeroPos - 1)
            Else
                RegGetstring = strBuf
            End If
        End If
    End If
    If RegGetstring = "" Then
    RegGetstring = "No! Value or Key Doesn't exist"
    End If
End Function

Function RegGetdword(ByVal hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    'EXAMPLE:
    '
    'text1.text = getdword(HKEY_CURRENT_USER
    '     , "Software\VBW\Registry", "Dword")
    '
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(hkey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            RegGetdword = lBuf
        End If
        '
    End If
    r = RegCloseKey(keyhand)
    If RegGetdword = "" Then
        RegGetdword = " No! Value or key doesn't exist"
    End If
    
End Function



Public Function RegDelKey(ByVal hkey As Long, ByVal strKey As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW")
    '
    Dim r As Long
    r = RegDeleteKey(hkey, strKey)
    If r = 0 Then
    
    DeleteKey = "Success"
    Else
    DeleteKey = "No! Key to Delete Or Key Not Found"
    End If
    
    
End Function

Public Function RegDelValue(ByVal hkey As Long, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    r = RegOpenKey(hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
    If r = 0 Then
    DeleteValue = "Success"
    Else
    DeleteValue = "No! Value to Delete"
    End If
End Function

    
Public Sub RegSavekey(hkey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub


Function RegSaveDword(ByVal hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    'EXAMPLE"
    '
    'Text1.text= SaveDword(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry", "Dword", text1.text)
    '
    '
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call e
    '     rrlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
    If r = 0 Then
    SaveDword = "Success"
    Else
    SaveDword = "No! Failed to save Value"
    End If
    
End Function


Public Function RegSavestring(hkey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'text1.text= savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(hkey, strPath, keyhand)
    r = RegCreateKey(hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
    If r = 0 Then
    
    RegSavestring = "Success"
    Else
    RegSavestring = "No! Key to Delete Or Key Not Found"
    End If
    
End Function

Public Function itsok(msg As String) As Boolean
If Left(itsok, 3) = "No!" Then
    itsok = 0
Else
    itsok = 1
End If
End Function

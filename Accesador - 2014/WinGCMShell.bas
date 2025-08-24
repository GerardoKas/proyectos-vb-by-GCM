Attribute VB_Name = "WinGCMShell"
Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long



'Display the title bar text of all top-level windows.  This
' task is given to the callback function, which will receive each handle individually.
' Note that if the window has no title bar text, it will not be displayed (for clarity's sake).

' *** Place this code in a module.  This is the callback function. ***
' This function displays the title bar text of the window identified by hwnd.
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
  Dim slength As Long, buffer As String  ' title bar text length and buffer
  Dim retval As Long  ' return value
  Static winnum As Integer  ' counter keeps track of how many windows have been enumerated

  winnum = winnum + 1  ' one more window enumerated....
  slength = GetWindowTextLength(hWnd) + 1  ' get length of title bar text
  If slength > 1 Then ' if return value refers to non-empty string
    buffer = Space(slength)  ' make room in the buffer
    retval = GetWindowText(hWnd, buffer, slength)  ' get title bar text
    'Debug.Print "Window #"; winnum; " : ";  ' display number of enumerated window
    venttext = Left(buffer, slength - 1) ' display title bar text of enumerated window
    frmVentanas.lstWindows.AddItem venttext
  End If

  EnumWindowsProc = 1  ' return value of 1 means continue enumeration
End Function

Sub ventanas()
' *** Place this code wherever you want to enumerate the windows. ***
Dim retval As Long  ' return value
frmVentanas.Visible = True
' Use the above callback function to list all of the enumerated windows.  Note that lParam is
' set to 0 because we don't need to pass any additional information to the function.
retval = EnumWindows(AddressOf EnumWindowsProc, 0)
End Sub

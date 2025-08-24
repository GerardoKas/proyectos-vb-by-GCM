Attribute VB_Name = "modConnectToWebSite"
' modConnectToWebSite
' 1998/06/10 Written by Dan Newsome, [newsomed@earthlink.net]
' 1998/06/10 Modified by Larry Rebich, [lrebich@earthlink.net]
' 2000/04/10 Copyright © 2000, Larry Rebich
' 2000/04/10 larry@buygold.net, www.buygold.net, 760.771.4730
' 2000/04/10 Use ShellExecute found at Randy Birch's [http://www.mvps.org/vbnet] site.

    Option Explicit
    DefLng A-Z
    
    Private Const SW_SHOWNORMAL As Long = 1

    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
         ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
    Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    '

Public Function ConnectToWebSite(frm As Form, sBrowserExec As String, sURL As String) As Boolean
    Const SW_SHOWNORMAL = 1
    Dim lRetVal As Long
    Dim sTemp   As String
    

    lRetVal = ShellExecute(frm.hWnd, "open", sBrowserExec, sURL, sTemp, SW_SHOWNORMAL)
'    lRetVal = ShellExecute(frm.hWnd, "open", sURL, "", sTemp, SW_SHOWNORMAL)           '1998/07/31 This works as well
    If lRetVal > 32 Then            ' OK
        ConnectToWebSite = True
    End If

End Function

Public Function GetBrowserExe() As String
    Dim sFileName   As String
    Dim sBrowserExec As String * 255
    Dim lRetVal     As Long
    Dim iFN         As Integer
    Dim sTemp       As String
    
    sBrowserExec = Space(255)
    sFileName = App.Path & "\temphtm.HTM"
    
    iFN = FreeFile()                    ' Get unused file number
    
    Open sFileName For Output As #iFN   ' Create temp HTML file
    Print #iFN, "<HTML> <\HTML>"        ' Output text
    Close #iFN                          ' Close file
    
    ' Then find the application associated with it.
    lRetVal = FindExecutable(sFileName, sTemp, sBrowserExec)
    ' If an application return the name
    If lRetVal <= 32 Or IsEmpty(sBrowserExec) Then ' Error
    Else
        GetBrowserExe = Trim$(sBrowserExec)
    End If
    Kill sFileName  ' delete temp HTML file

End Function


Public Function OpenURL(sURL As String) As Boolean
    Dim lRtn As Long
    lRtn = RunShellExecute("open", sURL, 0&, 0&, SW_SHOWNORMAL)
    If lRtn > 32 Then       'success
        OpenURL = True
    End If
End Function

Private Function RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long) As Long

    RunShellExecute = ShellExecute(GetDesktopWindow(), sTopic, sFile, sParams, sDirectory, nShowCmd)

End Function





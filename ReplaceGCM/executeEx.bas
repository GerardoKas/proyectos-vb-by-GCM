Attribute VB_Name = "executeEx"
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SW_SHOWNORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMINNOACTIVE = 7

Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As _
    SHELLEXECUTEINFO) As Long
Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal _
    dwMilliseconds As Long) As Long
Public Const INFINITE = &HFFFF
Public Const WAIT_TIMEOUT = &H102

' *** Place the following code inside window Form1. ***
Public Sub ejecutar(comando)
    Dim sei As SHELLEXECUTEINFO  ' structure used by the function
    Dim retval As Long  ' return value
    
    ' Load the information needed to open C:\Docs\readme.txt
    ' into the structure.
    With sei
        ' Size of the structure
        .cbSize = Len(sei)
        ' Use the optional hProcess element of the structure.
        .fMask = SEE_MASK_NOCLOSEPROCESS
        ' Handle to the window calling this function.
        .hwnd = Form2.hwnd
        ' The action to perform: open the file.
        .lpVerb = "open"
        ' The file to open.
        .lpFile = comando ' "C:\Docs\readme.txt"
        ' No additional parameters are needed here.
        .lpParameters = ""
        ' The default directory -- not really necessary in this case.
        '.lpDirectory = "C:\Docs\"
        ' Simply display the window.
        .nShow = SW_SHOWMINNOACTIVE
        ' The other elements of the structure are either not used
        ' or will be set when the function returns.
    End With
    
    ' Open the file using its associated program.
    retval = ShellExecuteEx(sei)
    If retval = 0 Then
        ' The function failed, so report the error.  Err.LastDllError
        ' could also be used instead, if you wish.
        Select Case sei.hInstApp
        Case SE_ERR_FNF
            Debug.Print "The file C:\Docs\readme.txt was not found."
        Case SE_NOASSOC
            Debug.Print "No program is associated with *.txt files."
        Case Else
            Debug.Print "An unexpected error occured."
        End Select
    Else
        ' Wait for the opened process to close before continuing.  Instead
        ' of waiting once for a time of INFINITE, this example repeatedly checks to see if the
        ' is still open.  This allows the DoEvents VB function to be called, preventing
        ' our program from appearing to lock up while it waits.
        Do
            DoEvents
            retval = WaitForSingleObject(sei.hProcess, 0)
        Loop While retval = WAIT_TIMEOUT
        Debug.Print "Notepad (or whatever program was opened) has just closed."
    End If
End Sub


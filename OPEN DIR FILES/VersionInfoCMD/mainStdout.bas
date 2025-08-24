Attribute VB_Name = "mainStdout"
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_ERROR_HANDLE = -12&

Private Declare Function GetStdHandle Lib "kernel32" (ByVal HandleType As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
       (ByVal hFile As Long, _
        ByVal lpBuffer As Any, _
        ByVal cToWrite As Long, _
        ByRef cWritten As Long, _
        Optional ByVal lpOverlapped As Long) As Long

'To keep things simple, wrap these two APIs in a single helper function like so:

Private Function ConsoleWrite(sText As String) As Long
    stdout = GetStdHandle(STD_OUTPUT_HANDLE)
    sText = sText '& Chr(0)
    r = WriteFile(stdout, ByVal sText, Len(sText), ConsoleWrite, 0)
    CloseHandle (stdout)
    Debug.Print r
    
End Function

'And now all you need to do is call your helper function when you want to write to the console:

Private Sub Main()
Dim file As String
If InStr(1, Command$(), "/VISUAL") > 0 Then
Form1.Visible = True
file = Replace(Command$, "/VISUAL", "")
file = Replace(Trim(file), Chr(34), "")
Form1.Text1 = file
Else
file = Command$()
file = Replace(file, """", "")

Con.Initialize
If file = "" Then
    ConsoleWrite "Uso:" & vbclrf & App.EXEName & " [/VISUAL] <exefilename>" & vbCrLf
Else
    Con.WriteLine Version1.Version1(file), True, conStandardOutput
End If
End If
End Sub





Public Sub Main_example()
   Dim sData As String
   Dim sMessage As String
'http://vb.mvps.org/samples/project.asp?id=console
   ' Required in all MConsole.bas supported apps!
   Con.Initialize

   ' Check to see if we have any waiting input.
   If Con.Piped Then
      ' Slurp it all in a single stream.
      sData = Con.ReadStream()
      ' Just to prove we did it, place text on clipboard.
      Clipboard.Clear
      Clipboard.SetText sData
      ' Write some debugging information.
      Con.DebugOutput "Wrote " & CStr(Len(sData)) & _
                      " characters to clipboard."
   Else
      sMessage = "No redirection detected; nothing to read?"
      ' Send error condition to Standard Error!
      Con.WriteLine sMessage, True, conStandardError
      Con.DebugOutput sMessage
      ' Set an exit code appropriate to this error.
      ' Application MUST BE COMPILED TO NATIVE CODE to
      ' avoid a GPF in the runtime!!!
      Con.ExitCode = 1
   End If
End Sub



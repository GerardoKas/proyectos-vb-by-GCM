VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CMD VIEWER"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H80000005&
   FillStyle       =   0  'Solid
   Icon            =   "frmSalida.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVETEXT"
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   5625
      Width           =   1230
   End
   Begin VB.ComboBox cmbTexto 
      Height          =   315
      ItemData        =   "frmSalida.frx":014A
      Left            =   0
      List            =   "frmSalida.frx":0157
      TabIndex        =   5
      Text            =   "cmd /c tree ""c:\archivos de programa"""
      ToolTipText     =   "LINEA DE COMANDOS"
      Top             =   0
      Width           =   3435
   End
   Begin VB.ComboBox cmbListaAscii 
      Height          =   315
      ItemData        =   "frmSalida.frx":018D
      Left            =   2295
      List            =   "frmSalida.frx":019A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Tablas Ascii"
      Top             =   360
      Width           =   1185
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      ItemData        =   "frmSalida.frx":01B8
      Left            =   3510
      List            =   "frmSalida.frx":01CB
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Fuente"
      Top             =   360
      Width           =   1140
   End
   Begin VB.ComboBox cmbChar 
      Height          =   315
      ItemData        =   "frmSalida.frx":0206
      Left            =   990
      List            =   "frmSalida.frx":0213
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Conversion de Caracter"
      Top             =   360
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXEC"
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   4560
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   675
      Width           =   4650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long


Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As Any, lpProcessInformation As Any) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWNORMAL = 1

Const STARTF_USESHOWWINDOW = &H1
Const INFINITE = -1&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&

Public TEXTOENTERO As String

Private Function ExecCmd(ByVal CmdLine As String) As String
    'Ejecuta el comando indicado, espera a que termine
    'y redirige la salida hacia VB

    Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long, hWritePipe As Long
    Dim bytesread As Long, mybuff As String
    Dim i As Integer
    
    Dim sReturnStr As String
    
    '=== Longitud de la cadena, en teoría 64 KB,
    '   pero no en la práctica
    'mybuff = String(64 * 1024, Chr$(65))
    '
    'sa=SECURITY DESCRIPTOR
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If ret = 0 Then
        '===Error
        ExecCmd = "Error: CreatePipe failed. " & Err.LastDllError
        Exit Function
    End If
    start.cb = Len(start)
    start.hStdOutput = hWritePipe
    start.dwFlags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    start.wShowWindow = SW_SHOWNORMAL
    
    ' Start the shelled application:
    ret& = CreateProcessA(0&, CmdLine$, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If ret <> 1 Then
        '===Error
        sReturnStr = "Error: CreateProcess failed. " & Err.LastDllError
    End If
    
    ' Wait for the shelled application to finish:
 On Error Resume Next
    mybuff = String(80 * 25, Chr$(65))
    ret = 1
    Do
    ret = WaitForSingleObject(proc.hProcess, 2 * 1000)
    If ret < 0 Then Exit Function
    bSuccess = ReadFile(hReadPipe, mybuff, Len(mybuff), bytesread, 0&)
    If bSuccess = 1 Then
        sReturnStr = sReturnStr & Left(mybuff, bytesread)
    Else
        '===Error
        sReturnStr = "Error: ReadFile failed. " & Err.LastDllError
        Exit Function
    End If
    Loop While ret
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    ret = CloseHandle(hWritePipe)
    
    ExecCmd = sReturnStr
    TEXTOENTERO = sReturnStr
End Function


Private Sub cmbFont_Click()
If cmbFont.ListIndex = 0 Then Exit Sub
Text1.Font = cmbFont.List(cmbFont.ListIndex)
Text1.FontSize = 10
Text1.FontBold = False
End Sub

Private Sub cmbListaAscii_Click()
Select Case cmbListaAscii.ListIndex
Case 1
AscList1
Case 2
Asclist2
End Select
TEXTOENTERO = Text1.Text
End Sub


Private Sub cmbTexto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub cmdSAVE_Click()
If Len(TEXTOENTERO) = 0 Then
    MsgBox "No hay nada ejecutado"
    Exit Sub
End If
mpath = App.Path
If Right(mpath, 1) <> "\" Then mpath = mpath & "\"
mfile = mpath & "SavedCommand" & Format(Now, "dd-mm-yyyy_hh-mm") & ".txt"
f = FreeFile
Open mfile For Output As #f
Print #f, TEXTOENTERO
Close #f
MsgBox "GUARDADO EN " & vbCrLf & mfile
End Sub

Private Sub Command1_Click()
    Text1.Text = ExecCmd(cmbTexto)
End Sub


Private Sub AscList1()
Dim str As String
For i = 1 To 255
str = str & "char(" & i & ") = (" & Chr(i) & ") " & vbCrLf
Next
Text1 = str
End Sub
Private Sub Asclist2()
Dim str As String
For i = 1 To 1024
str = str & "(" & i & ")=" & ChrW(i) & vbTab
If i Mod 4 = 0 Then
str = str & vbCrLf
End If
Next
Text1 = str
End Sub



Private Sub Form_Load()
cmbFont.ListIndex = 0
cmbChar.ListIndex = 0
cmbListaAscii.ListIndex = 0
Show
cmbTexto.SetFocus
'Asigna al TextBox lo que se introduzca en la línea de comandos
'Text1 = ExecCmdPipe(Command$)
End Sub
 

Private Sub cmbChar_Click()
If cmbChar.Text = "DOS2WIN" Then
DOS2WIN
ElseIf cmbChar.Text = "WIN2DOS" Then
WIN2DOS
End If
End Sub


Private Sub DOS2WIN()
Dim cad1 As String, cad2 As String
cad1 = Text1
cad2 = String$(Len(Text1) + 1, vbNullChar)
r = OemToChar(cad1, cad2)
'MsgBox r
Text1 = delNull(cad2)
End Sub

Private Sub WIN2DOS()
Dim cad1 As String, cad2 As String
cad1 = Trim(Text1)
cad2 = String$(Len(Text1) + 1, vbNullChar)
r = CharToOem(cad1, cad2)
'MsgBox r
Text1 = delNull(cad2)
End Sub

Function delNull(cad) As String
delNull = Left$(cad, InStr(1, cad, Chr(0)) - 1)
End Function

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
cmbTexto.Width = Me.ScaleWidth - 10
Text1.Width = Me.ScaleWidth
Text1.Height = Me.ScaleHeight - Text1.Top - cmdSAVE.Height
cmdSAVE.Top = Me.ScaleHeight - cmdSAVE.Height
End Sub

Private Sub Text1_Change()
Text1.Font.Charset = 1
End Sub


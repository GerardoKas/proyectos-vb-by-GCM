VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.DriveListBox Unidad 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   4320
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Result1 
      BackColor       =   &H80000018&
      Caption         =   "Resultados."
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function Musica(ByVal Archivo As String, ByVal modo As Long) As Long
Err = sndPlaySound(Archivo, SND_ASYNC + modo)

End Function
Function FindFilesAPI(path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer)

Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path

Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer

If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)

If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If hSearch = FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
End If
End If
Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If

' Walk through this directory and sum file sizes.

hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True

If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
MAXDWORD) + WFD.nFileSizeLow

FileCount = FileCount + 1
List1.AddItem path & FileName
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If

' If there are sub-directories...

If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(path & _
dirNames(i) & "\", SearchStr, FileCount, DirCount)
Next i
End If

End Function
Private Sub DIR1_Change()
Dir1.path = Unidad

Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
List1.Clear
SearchPath = Dir1
FindStr = "*.MP3"

FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
List1 = NumDirs
Text3.Text = Dir1
'NumFiles & " Files found in " & NumDirs + 1 & _
" Directories"
Text4.Text = "Size of files found under " & SearchPath & " = " & _
Format(FileSize, "#,###,###,##0") & " Bytes"
Screen.MousePointer = vbDefault

End Sub

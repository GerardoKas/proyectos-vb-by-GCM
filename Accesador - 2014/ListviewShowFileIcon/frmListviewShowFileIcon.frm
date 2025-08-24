VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListviewShowFileIcon 
   Caption         =   "frmListviewShowFileIcon"
   ClientHeight    =   5610
   ClientLeft      =   2865
   ClientTop       =   2895
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   9195
   Begin VB.CommandButton Command1 
      Caption         =   "&Browse for Folder"
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Browse"
      Top             =   3600
      Width           =   3075
   End
   Begin VB.PictureBox picLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   780
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   600
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListviewShowFileIcon.frx":0000
            Key             =   "Dummy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   1380
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListviewShowFileIcon.frx":015A
            Key             =   "Dummy"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cycle ListView Views"
      Height          =   495
      Index           =   0
      Left            =   5520
      TabIndex        =   0
      ToolTipText     =   "Cycle Views"
      Top             =   3600
      Width           =   3075
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   5741
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Loading..."
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Info (via the web)"
         Index           =   0
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Other Tips (via the web)"
         Index           =   1
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&About"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmListviewShowFileIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmListviewShowFileIcon
' 2003/03/05 Copyright © 2003, Larry Rebich, using the DELL7500
' 2003/03/05 larry@larryrebich.com, www.larryrebich.com, 760-771-4730

' See Also from Brad Martinez
' http://www.mvps.org/btmtz/listview/

    Option Explicit
    DefLng A-Z
    '
    Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    '
    
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0  'cycle views
            DoCycleViews
        Case 1  'browse for folder
            DoBrowse
    End Select
End Sub

Private Sub DoBrowse()
' See: http://www.buygold.net/v04n06/v04n06.html
    Dim o As New cBrowseFolder
    With o
        .lhWnd = Me.hWnd
        .sTitle = "Select a Folder"
        .sFolder = Me.Caption       'last one stored in the caption - easy
        .lFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE
        .ShowBrowse
        If Not .bCancelled Then
            Me.Caption = .sFolder   'store in the form's caption
            LoadListView .sFolder   'load the listview
        End If
    End With
End Sub

Private Sub DoCycleViews()
    Dim iView As Integer
    iView = Me.ListView1.View
    iView = iView + 1
    iView = iView Mod 4
    With Me.ListView1
        .View = iView
        .Refresh        'try to fix bug in listview showing icons in wrong places
        .View = iView
    End With
End Sub

Private Sub Form_Load()
    SetupForm
    Show            'show the form
    DoEvents        'allow form to refresh
    LoadListView    'load the WinDir
    Me.ListView1.Visible = True
End Sub

Private Sub SetupForm()
    With Me
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2 'simple center
    End With
    SetupListView
End Sub

Private Sub SetupListView()
    With Me.ListView1
        .View = lvwReport           'start in report view
        .HideSelection = False      'don't hide the selection
        .GridLines = True
        .LabelEdit = lvwManual      'no label editing
        ' add some columns
        .ColumnHeaders.Add , "name", "Name", 4000
        .ColumnHeaders.Add , "size", "Size", 1200, lvwColumnRight
        .ColumnHeaders.Add , "type", "Type", 3000
    End With
End Sub

Private Sub LoadListView(Optional vntFolder As Variant)
    Dim itm As ListItem
    Dim sLoadFolder As String
    Dim sFolder As String
    Dim arySomeFiles() As String
    Dim iSomeFilesCount As Integer
    Dim sFile As String
    Dim i As Integer
    Dim lFlag As Long
    Const sKeyPrefix As String = "@"    '2003/11/30 Needed in case file name is all numeric.
    
    If Not IsMissing(vntFolder) Then    'if user browses then folder is supplied
        sFolder = vntFolder
    End If
    
    If sFolder = "" Then                'if none then use WinDir
        sLoadFolder = Environ("WinDir") 'environ way
        sLoadFolder = GetWindowsDir()   'api way
    Else
        sLoadFolder = sFolder           'user supplied
    End If
    
    If Right$(sLoadFolder, 1) <> "\" Then   'add a right backslash if none
        sLoadFolder = sLoadFolder & "\"     'add it
    End If
    Me.Caption = sLoadFolder                'show in caption
    Me.Label1.Caption = "Loading " & sLoadFolder & "..."    'loading shows in label
    
    lFlag = vbDirectory + vbHidden + vbSystem   'get all directory types
    
    On Error GoTo LoadListViewEH
    sFile = Dir$(sLoadFolder & "*.*", lFlag)    'get the first file
    If sFile <> "" Then                         'any file?
        iSomeFilesCount = iSomeFilesCount + 1   'have a file so increment the files counter
        ReDim Preserve arySomeFiles(1 To iSomeFilesCount)   'redim the array
        arySomeFiles(iSomeFilesCount) = sFile   'store the first file in the array
        While sFile <> ""                       'any more files
            sFile = Dir$()                      'get next file
            If sFile <> "" Then                 'any file
                iSomeFilesCount = iSomeFilesCount + 1   'increment the files counter
                ReDim Preserve arySomeFiles(1 To iSomeFilesCount)   'redim the array
                arySomeFiles(iSomeFilesCount) = sFile   'store the file in the array
            End If
        Wend
    End If
    Me.ListView1.ListItems.Clear    'clear the listview
    Me.ListView1.Visible = False    'faster loading if not 'refreshed' with each addition
    LV_ClearImageListsAndRelinkToListView Me.ListView1, Me.imgLarge, Me.imgSmall  'make sure the image lists are clear of icons
    With Me.ListView1
        For i = 1 To iSomeFilesCount        'for each file add it to the listview
            If arySomeFiles(i) = "." Or arySomeFiles(i) = ".." Then 'don't show 'up' folders
            Else
                ' The next line does all the work
                LV_AddFileToListView sLoadFolder, arySomeFiles(i), "@", Me.ListView1, 2, Me.imgLarge, Me.imgSmall, Me.picLarge, Me.picSmall
                Set itm = .ListItems("@" & arySomeFiles(i))   'used for next statement to get the file length
                itm.SubItems(1) = FormatLength(FileLen(sLoadFolder & arySomeFiles(i)))
            End If
        Next
    End With
    
    If Me.ListView1.ListItems.Count > 0 Then    'any files
        Me.ListView1.Visible = True
    Else
        Me.ListView1.Visible = False            'no files
        Me.Label1.Caption = "No files in " & sLoadFolder    'tell 'em no files
    End If
    Exit Sub
LoadListViewEH:
    Me.Label1.Caption = Err.Description & " " & sLoadFolder 'show error description in label
    Me.ListView1.Visible = False                            'hide listview
End Sub

Private Function FormatLength(lLength As Long) As String
    Dim sLength As String
    sLength = Format$(lLength, "###,###,###,###")
    FormatLength = Right$(Space(12) & sLength, 12)  'pad with left spaces to assist sorting
End Function

Private Function GetWindowsDir() As String
'From: http://www.desertware.com/vbuniverse/getwindowsdirectories.html

    Dim sDir As String * 255
    Dim lReturn As Long
    Dim lSize As Long
    
    sDir = Space$(255)
    lSize = Len(sDir)
    lReturn = GetWindowsDirectory(sDir, lSize)
    
    If lReturn > 0 Then
        GetWindowsDir = Mid$(sDir, 1, lReturn)
    Else
        GetWindowsDir = Left$(sDir, InStr(1, sDir, vbNullChar) - 1)
    End If

End Function

Private Sub Form_Resize()
' form is being resized - move the controls
    Dim l, t, w, h
    If Me.WindowState <> vbMinimized Then
        With Me.Command1(0)
            l = (Me.ScaleWidth \ 2) + 30
            t = Me.ScaleHeight - .Height - 90
            .Move l, t
        End With
        With Me.Command1(1)
            l = (Me.ScaleWidth \ 2) - .Width - 30
            .Move l, t
        End With
        With Me.ListView1
            l = 30
            t = l
            w = Me.ScaleWidth - l * 2
            h = Me.ScaleHeight - (t * 4) - Me.Command1(0).Height - 60
            .Move l, t, w, h
        End With
        With Me.Label1
            l = (Me.ScaleWidth - .Width) \ 2
            t = ((Me.ScaleHeight - .Height) \ 2) * 0.8
            .Move l, t
        End With
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' sort them
    With Me.ListView1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .SortKey = ColumnHeader.Index - 1
        .Sorted = True
    End With
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Select Case Index
        Case 0  'exit
            Unload Me
    End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0  'info
            DoInfo Me, 6, 9
        Case 1  'other tips
            DoTips
        Case 2  'about
            DoAbout Me
    End Select
End Sub

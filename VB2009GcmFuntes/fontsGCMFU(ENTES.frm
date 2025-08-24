VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4704
   ClientLeft      =   64
   ClientTop       =   336
   ClientWidth     =   7072
   LinkTopic       =   "Form1"
   ScaleHeight     =   4704
   ScaleWidth      =   7072
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "GUARDAR"
      Height          =   400
      Left            =   5376
      TabIndex        =   10
      Top             =   1280
      Width           =   1552
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IMPRIMIR"
      Height          =   400
      Left            =   5376
      TabIndex        =   8
      Top             =   768
      Width           =   1552
   End
   Begin MSComDlg.CommonDialog cmdlg1 
      Left            =   3328
      Top             =   2048
      _ExtentX        =   945
      _ExtentY        =   945
      _Version        =   327680
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Texto ASCII"
      Height          =   400
      Index           =   1
      Left            =   5376
      TabIndex        =   7
      Top             =   384
      Width           =   1680
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Texto Ejemplo"
      Height          =   272
      Index           =   0
      Left            =   5376
      TabIndex        =   6
      Top             =   128
      Width           =   1552
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Subrayado"
      Height          =   400
      Left            =   3712
      TabIndex        =   5
      Top             =   1152
      Width           =   1168
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Negrita"
      Height          =   272
      Left            =   3712
      TabIndex        =   4
      Top             =   640
      Width           =   1296
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Italicas"
      Height          =   272
      Left            =   3712
      TabIndex        =   3
      Top             =   128
      Width           =   1296
   End
   Begin VB.ListBox lstSizes 
      Height          =   1728
      ItemData        =   "fontsGCMFU(ENTES.frx":0000
      Left            =   2432
      List            =   "fontsGCMFU(ENTES.frx":0022
      TabIndex        =   2
      Top             =   128
      Width           =   1168
   End
   Begin VB.ListBox List1 
      Height          =   4432
      Left            =   0
      TabIndex        =   1
      Top             =   128
      Width           =   2320
   End
   Begin VB.TextBox Text1 
      Height          =   2608
      Left            =   2432
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "fontsGCMFU(ENTES.frx":004D
      Top             =   1920
      Width           =   4624
   End
   Begin RichTextLib.RichTextBox rich1 
      Height          =   912
      Left            =   5248
      TabIndex        =   9
      Top             =   3584
      Width           =   1680
      _ExtentX        =   3101
      _ExtentY        =   1683
      _Version        =   327680
      Enabled         =   -1  'True
      TextRTF         =   $"fontsGCMFU(ENTES.frx":007D
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Text1.FontItalic = Check1.Value
End Sub

Private Sub Check2_Click()
Text1.FontBold = Check2.Value
End Sub

Private Sub Check3_Click()
Text1.FontUnderline = Check3.Value
End Sub

Private Sub Command1_Click()
cmdlg1.Color = 0
cmdlg1.DialogTitle = "IMPRIMIR"
cmdlg1.ShowPrinter
If cmdlg1.CancelError = False Then MsgBox "NO se imprimira": Exit Sub
Printer.Print "LISTADO DE FUEN TES " & Date
For i = 0 To List1.ListCount - 1
Printer.Font = List1.List(i)
'Printer.FontSize = CInt(lstSizes.List(lstSizes.ListIndex))
Printer.Print Text1.Text & vbCrLf
Next
Printer.EndDoc
End Sub

Private Sub Command2_Click()
'rich1.Text = Text1.Text
Dim a As String
a = Text1.Text
rich1.Font.Size = lstSizes.List(lstSizes.ListIndex)
rich1.TextRTF = "Listado de Fuentes del Sistema " & Date & vbCrLf
For i = 0 To List1.ListCount - 1
rich1.SelFontName = List1.List(i)
rich1.SelText = a & vbCrLf
Next
rich1.SaveFile "C:\FUENTES2011.rtf"
yes = MsgBox("Guardado en C:\FUENTES2011.rtf", vbYesNo, "Creado, Desea Abrirlo?")
If yes = vbYes Then
Shell "write C:\FUENTES2011.rtf"
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To Screen.FontCount - 1
List1.AddItem Screen.Fonts(i)
Next
'Text1.Font.Name = Screen.Fonts(0)
List1.Selected(0) = True
lstSizes.Selected(0) = True
x = GetSetting("ME", "DATA", "TEXTO", "El Veloz Kiwi Comia feliz Cardillo y Nieve")
If x <> "" Then
Text1.Text = x
Else
x = "El Veraz Veloz Feriz Comedia Cardillo y Kiewi"
End If
Text1.Text = x
lstSizes.ListIndex = 2
End Sub


Private Sub Form_Resize()
On Error Resume Next
List1.Height = Me.ScaleHeight - List1.Top
Text1.Height = Me.ScaleHeight - Text1.Top
Text1.Width = Me.ScaleWidth - Text1.Left

End Sub

Private Sub Form_Unload(Cancel As Integer)
Command1_Click
SaveSetting "ME", "DATA", "TEXTO", Text1.Text
End Sub

Private Sub List1_Click()
Text1.Font.Name = Screen.Fonts(List1.ListIndex)
'Text1.Font.Size = List1.List(List1.ListIndex)
Text1.Font.Bold = Check1.Value
Text1.FontUnderline = Check2.Value
Text1.Font.Italic = Check3.Value
End Sub

Private Sub lstSizes_Click()
Text1.FontSize = lstSizes.List(lstSizes.ListIndex)
End Sub

Private Sub Option1_Click(Index As Integer)
If Index = 0 Then
Text1.Text = "El Veraz Veloz Feriz Comedia Cardillo y Kiewi"
ElseIf Index = 1 Then
Text1.Text = ascii()
End If
End Sub
Private Function ascii()
t = ""
For i = 1 To 255
t = t & Chr(i)
Next
ascii = t
End Function

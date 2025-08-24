VERSION 5.00
Begin VB.Form frmCSS 
   Caption         =   "Form2"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form2"
   ScaleHeight     =   4875
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   0
      TabIndex        =   6
      Top             =   2430
      Width           =   5820
      Begin VB.CommandButton cmdLoadAttribs 
         Caption         =   "Load"
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   1935
         Width           =   870
      End
      Begin VB.ListBox lstAttribs 
         Height          =   2010
         Left            =   1620
         TabIndex        =   8
         Top             =   225
         Width           =   2625
      End
      Begin VB.ComboBox cmbAttribClases 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Disable"
      Height          =   285
      Left            =   1530
      TabIndex        =   5
      Top             =   45
      Width           =   1050
   End
   Begin VB.ComboBox cmbElements 
      Height          =   315
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   45
      Width           =   1410
   End
   Begin VB.TextBox txtCSS 
      Height          =   2040
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "frmEditar.frx":0000
      Top             =   405
      Width           =   5775
   End
   Begin VB.ComboBox cmbHojas 
      Height          =   315
      Left            =   45
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   1455
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar CSS"
      Height          =   285
      Left            =   4545
      TabIndex        =   0
      Top             =   45
      Width           =   1275
   End
   Begin VB.Label lblObjeto 
      AutoSize        =   -1  'True
      Caption         =   "OBJETO ACTUAL"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   2520
      Width           =   1305
   End
End
Attribute VB_Name = "frmCSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myObjCss As IHTMLStyle
Dim myObjStylesheet As IHTMLStyleSheet
Dim viendoSS As Boolean
Private Sub Check1_Click()
On Error Resume Next
    If Check1.Value = 1 Then
        myObjStylesheet.disabled = True
    Else
        myObjStylesheet.disabled = False
    End If
End Sub

Private Sub cmbAttribClases_Click()
lstAttribs.Clear
ReadWrite.setIniFile "CSSATTRIBS.INI"
mclase$ = ReadWrite.ReadFromFile("INIT", cmbAttribClases.Text)

Keys = ReadWrite.ReadKeys(mclase)
mkeys = Split(Keys, Chr(0))
For Each i In mkeys
    lstAttribs.AddItem i
Next

End Sub

Private Sub cmbElements_Click()
If cmbElements.ListIndex <= 0 Then Exit Sub
Check1.Enabled = False
cmbHojas.ListIndex = 0
t = cmbElements.Text
pts = Split(t, ",")
objid = pts(2)
Set myObjCss = myDoc.getElementById(objid).style
txtCSS = myObjCss.cssText
lblObjeto = cmbElements.Text
viendoSS = False
End Sub

Private Sub cmbHojas_Click()
num = cmbHojas.ListIndex
If num <= 0 Then Exit Sub
cmbElements.ListIndex = 0
Set myObjStylesheet = myDoc.styleSheets(num - 1)
txtCSS = myObjStylesheet.cssText
lblObjeto = cmbHojas.Text
If myObjStylesheet.disabled = True Then Check1.Value = 1 Else Check1.Value = 0
Check1.Enabled = True
viendoSS = True
End Sub

Private Sub cmdActualizar_Click()
If viendoSS = True Then
    myObjStylesheet.cssText = txtCSS
Else
    myObjCss.cssText = txtCSS
End If
End Sub

Private Sub cmdLoadAttribs_Click()
ReadWrite.setIniFile "CSSATTRIBS.INI"
Keys = ReadWrite.ReadKeys("INIT")
mkeys = Split(Keys, Chr(0))
For Each i In mkeys
    If i <> "" Then
        cmbAttribClases.AddItem i
    End If
Next
End Sub

Private Sub Form_Load()
Dim elem As IHTMLElement
n = 0
cmbHojas.AddItem "StyleSheets"
cmbElements.AddItem "Id Unicos"
cmbHojas.ListIndex = 0
cmbElements.ListIndex = 0

For Each i In myDoc.styleSheets
    If i.href = "" Then
        nm = "Hoja " & n
    Else
        nm = i.href
    End If
    cmbHojas.AddItem nm
    n = n + 1
Next

For Each elem In myDoc.All
    If elem.style.cssText <> "" Then
        cmbElements.AddItem elem.tagName & "," & elem.Id & "," & elem.uniqueID
    End If
Next
End Sub

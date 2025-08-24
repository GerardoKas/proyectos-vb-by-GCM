VERSION 5.00
Object = "{6446818E-FC0F-4E35-9697-05A0AAB43FD0}#18.0#0"; "GCMPicControl.ocx"
Begin VB.Form frmCrop 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "frmCrop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin PictureControl_GCM.GCM_Picture gcm1 
      Height          =   2760
      Left            =   45
      TabIndex        =   11
      Top             =   270
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4868
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   0
      TabIndex        =   0
      Top             =   3105
      Width           =   3885
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1215
         TabIndex        =   15
         Text            =   "0"
         Top             =   405
         Width           =   465
      End
      Begin VB.CommandButton cmdSpecial 
         Caption         =   "SPECIAL"
         Height          =   285
         Left            =   2835
         TabIndex        =   13
         Top             =   1035
         Width           =   915
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1710
         TabIndex        =   12
         Text            =   "0"
         Top             =   405
         Width           =   420
      End
      Begin VB.CheckBox chkAutoRecarga 
         Caption         =   "AutoRecarga"
         Height          =   285
         Left            =   2430
         TabIndex        =   10
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox chkDebug 
         Caption         =   "Debug"
         Height          =   195
         Left            =   2430
         TabIndex        =   9
         Top             =   495
         Width           =   1005
      End
      Begin VB.CommandButton cmdSaveAs 
         Caption         =   "Guardar Recorte..."
         Height          =   285
         Left            =   1125
         TabIndex        =   8
         Top             =   1035
         Width           =   1635
      End
      Begin VB.CommandButton cmdRestaurar 
         Caption         =   "Restaurar"
         Height          =   285
         Left            =   45
         TabIndex        =   7
         Top             =   1035
         Width           =   1005
      End
      Begin VB.CheckBox chkBackup 
         Caption         =   "Backup"
         Height          =   195
         Left            =   2430
         TabIndex        =   6
         Top             =   225
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "Aplicar"
         Height          =   285
         Left            =   45
         TabIndex        =   5
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lbl1 
         Caption         =   "x y"
         Height          =   240
         Left            =   90
         TabIndex        =   4
         Top             =   450
         Width           =   1140
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "w h"
         Height          =   195
         Left            =   1215
         TabIndex        =   3
         Top             =   765
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "PosX Posy"
         Height          =   240
         Left            =   45
         TabIndex        =   2
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Width Height"
         Height          =   195
         Left            =   1215
         TabIndex        =   1
         Top             =   180
         Width           =   930
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TAMANIO"
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   45
      Width           =   735
   End
End
Attribute VB_Name = "frmCrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posible As Boolean
Dim doShow As Boolean
Dim doBackup As Boolean
Private myX As Long, myY As Long, myW As Long, myH As Long

Private Sub chkAutoRecarga_Click()
If chkAutoRecarga.Value = 1 Then
    doShow = True
Else
    doShow = False
End If
End Sub

Private Sub chkBackup_Click()
If chkBackup.Value = 1 Then
    doBackup = True
Else
    doBackup = False
End If
End Sub

Private Sub chkDebug_Click()
If chkDebug.Value = 1 Then
    doDebug = True
    gcm1.doDebug = True
Else
    doDebug = False
    gcm1.doDebug = False
End If
End Sub

Private Sub cmdAplicar_Click()
If doBackup = False Then
    If MsgBox("Si no haces backup se perderá el original," & vbCrLf & "Deseas hacer Backup?", vbYesNo) = vbYes Then
        doBackup = True
    Else
        cmdRestaurar.Enabled = False
    End If
End If
    
If doBackup = True Then
    cmdRestaurar.Enabled = True
    myBackup = getBackupName(myFile)
    BackupIt myFile, myBackup 'la nueva conserva el mismo nombre
End If
recortar myX, myY, myW, myH
gcm1.LoadNew myFile

End Sub

Private Sub cmdRestaurar_Click()
If myBackup = "" Then MsgBox "Lo siento no has hecho backup. Se ha perdido el original": Exit Sub
If Dir$(myBackup) <> "" Then
restaura myFile, myBackup
gcm1.LoadNew myFile
End If
End Sub

Private Sub cmdSaveAs_Click()
myPath$ = basepath(myFile)
myname$ = basename(myFile)
myRecorte = SaveDialog(Me, "Archivos JPG|*.jpg", "Guardar Recorte", myPath, myname)
If myRecorte = "" Then Exit Sub
If LCase(Right(myRecorte, 4)) <> ".jpg" Then
    myRecorte = myRecorte & ".jpg"
End If
Debug.Print myRecorte
copiarRecorte myX, myY, myW, myH
gcm1.LoadNew myRecorte
End Sub

Private Sub cmdSpecial_Click()
frmParams.Show 0, Me
End Sub

Private Sub Form_Load()
cmd$ = Command$()
If cmd <> "" Then
myFile = sinComillas(cmd$)
gcm1.LoadNew myFile
End If
doBackup = chkBackup.Value
doShow = chkAutoRecarga.Value
doDebug = chkDebug.Value
gcm1.doDebug = chkDebug.Value
End Sub

Private Sub Form_Resize()
On Error Resume Next
gcm1.Width = Me.ScaleWidth - 40
gcm1.Height = Me.ScaleHeight - Frame1.Height - gcm1.top - 40
Frame1.top = gcm1.Height + gcm1.top
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmParams
Unload frmCrop
End
End Sub

Private Sub gcm1_changedMark(X As Variant, Y As Variant, w As Variant, h As Variant)
lbl1 = X & " - " & Y
'lbl2 = w & " x " & h
If w > 2 And h > 2 Then
    posible = True
    myX = X
    myY = Y
    myW = w
    myH = h
    frmParams.ColocarDatos gcm1
Else
    posible = False
End If
End Sub

Private Sub gcm1_newFileLoaded(filename As Variant)
myFile = filename
gcm1.zoomAjustar
'gcm1.Marcar 0, 0, 200, 200
Label3.Caption = gcm1.imagePixelWidth & "x" & gcm1.imagePixelHeight
End Sub

Private Sub gcm1_zArrastrando(w As Variant, h As Variant)
txtWidth = w
txtHeight = h
End Sub

Private Sub gcm1_zIniciadaMarka(X As Variant, Y As Variant)
lbl1 = X & " - " & Y & "..."
End Sub


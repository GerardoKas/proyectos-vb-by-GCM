Attribute VB_Name = "MCIprg"

Public Sub abrefile(archivo As String, hwnd As Long)
Dim MciAbre As MCI_OVLY_OPEN_PARMS
Dim MCIrect As MCI_OVLY_RECT_PARMS

Dim val As Long
MciAbre.lpstrElementName = archivo
MciAbre.dwStyle = WS_CHILD
MciAbre.hWndParent = hwnd
MciAbre.lpstrAlias = "z1"
val = MCIOPEN(0, MCI_OPEN, MCI_OPEN_ELEMENT Or MCI_OVLY_OPEN_PARENT, MciAbre)
If (val) Then errorMCI (val)
val = MCIRECTANGULO(devid, MCI_WHERE, MCI_OVLY_WHERE_SOURCE, MCIrect)
If (val) Then errorMCI val








End Sub
Private Sub errorMCI(numero As Long)
Dim error As String
error = Space(1000)

    MsgBox mciGetErrorString(numero, error, 1000)
End Sub

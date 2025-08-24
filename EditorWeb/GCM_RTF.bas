Attribute VB_Name = "GCM_RTF"
Public rtf As RichTextBox
Private mPalabra As String

Sub colorTAGS()
Text = rtf.Text
rtf.SelStart = 0
rtf.SelLength = Len(rtf.Text)
rtf.SelColor = vbBlack
rtf.Visible = False
Do While rtf.SelStart < Len(rtf.Text)
rtf.span "<", True, False
rtf.SelLength = 1
rtf.SelColor = RGB(255, 0, 0)
rtf.SelStart = rtf.SelStart + 1
rtf.SelLength = 0
Loop
End Sub

Sub searchWord(palabra As String)
mPalabra = palabra
pos = rtf.Find(palabra, 0)
rtf.SelLength = Len(palabra)
End Sub

Sub searchNext()
pos = rtf.Find(mPalabra, rtf.SelStart + 1)
rtf.SelLength = Len(mPalabra)
End Sub

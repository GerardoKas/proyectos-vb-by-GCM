VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "simpleObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim thisObj As String
Dim props As Collection


Public Sub set_prop(name As String, value As String)
props.Add value, name
End Sub

Public Property Get prop(name As String)
If Not props.Item(name) Is Error Then
    prop = props.Item(name)
End If
End Property


Public Sub alCambiar(propName, metodo)

End Sub

Private Sub Class_Initialize()
Set props = New Collection
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oObject As Object
Private sMethod As String
Private bInstantiated As Boolean

Public Sub Class_Initialize()
    bInstantiated = False
End Sub

Public Sub Set_Destination(oInObject As Object, sInMethod As String)
Attribute Set_Destination.VB_Description = "el objeto donde esta el webbrowser y y el metodo que sera usado"

    Set oObject = oInObject
    sMethod = sInMethod
    bInstantiated = True
  
End Sub

Public Sub Default_Method()
Attribute Default_Method.VB_Description = "este obj debe ser asociado al document.onclick o cualquier evento que quieras"
Attribute Default_Method.VB_UserMemId = 0
'MsgBox "por fin"
    If bInstantiated Then
        'por ahora no me fue con el web1 solo con form1
        CallByName oObject, sMethod, VbMethod
    Else
        MsgBox "Je. Debes iniciar antes el objecto indicando el metodo para usar"
    End If
    
End Sub
'sacado de msdn reusing web browser

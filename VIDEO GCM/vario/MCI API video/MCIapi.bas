Attribute VB_Name = "MCIapi"

Public Const MCI_ANIM_PLAY_FAST = &H40000
Public Const MCI_ANIM_PLAY_REVERSE = &H20000
Public Const MCI_ANIM_PLAY_SCAN = &H100000
Public Const MCI_ANIM_PLAY_SLOW = &H80000
Public Const MCI_ANIM_PLAY_SPEED = &H10000
Public Const MCI_ANIM_WHERE_DESTINATION = &H40000
Public Const MCI_ANIM_WHERE_SOURCE = &H20000
Public Const MCI_CUT = &H851
Public Const MCI_COPY = &H852
Public Const MCI_CLOSE = &H804
Public Const MCI_CUE = &H830
Public Const MCI_DEVTYPE_OVERLAY = 515
Public Const MCI_FORMAT_MILLISECONDS = 0
Public Const MCI_LOAD = &H850
Public Const MCI_OPEN = &H803

Public Const MCI_OVLY_WHERE_DESTINATION = &H40000
Public Const MCI_OVLY_WHERE_SOURCE = &H20000
Public Const MCI_OVLY_WINDOW_HWND = &H10000

Public Const MCI_OPEN_ELEMENT = &H200&

Public Const MCI_OPEN_ALIAS = &H400&
Public Const MCI_OVLY_OPEN_PARENT = &H20000
Public Const MCI_OVLY_OPEN_WS = &H10000

Public Const MCI_WHERE = &H843
Public Const MCI_STOP = &H808
Public Const MCI_SET_VIDEO = &H1000&
Public Const MCI_PLAY = &H806
Public Const MCI_PUT = &H842
'Tipos
Type MCI_OVLY_RECT_PARMS
        dwCallback As Long
        rc As RECT
End Type
Type MCI_OVLY_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwStyle As Long
        hWndParent As Long
End Type
'La ventana tipo
Public Const WS_CHILD = &H40000000



'DECLARACIONES'
Declare Function MCIRECTANGULO Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, _
ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As MCI_OVLY_RECT_PARMS) As Long
Declare Function MCIOPEN Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, _
ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As MCI_OVLY_OPEN_PARMS) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long


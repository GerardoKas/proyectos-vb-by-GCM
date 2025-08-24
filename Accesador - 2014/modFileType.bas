Attribute VB_Name = "modFileType"
Private iconsIndex As Integer
'Private fileIcon As stdpicture
Private fileTypeString As String
Private fileExtension As String
Private file As String
Private myPic As PictureBox
Private myImageList As imagelist
Private fileTypes As Collection
Private iniciado As Boolean

    Const MAX_PATH = 260

    Const SHGFI_DISPLAYNAME = &H200
    Const SHGFI_EXETYPE = &H2000
    Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
    Const SHGFI_LARGEICON = &H0        ' Large icon
    Const SHGFI_SMALLICON = &H1        ' Small icon
    Const ILD_TRANSPARENT = &H1        ' Display transparent
    Const SHGFI_SHELLICONSIZE = &H4
    Const SHGFI_TYPENAME = &H400
    Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

    Public Type SHFILEINFO
        hIcon As Long
        iIcon As Long
        dwAttributes As Long
        szDisplayName As String * MAX_PATH
        szTypeName As String * 80
    End Type
    
    Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
       (ByVal pszPath As String, _
        ByVal dwFileAttributes As Long, _
        psfi As SHFILEINFO, _
        ByVal cbSizeFileInfo As Long, _
        ByVal uFlags As Long) As Long
    
    Private Declare Function ImageList_Draw Lib "comctl32.dll" _
       (ByVal himl&, ByVal i&, ByVal hDCDest&, _
        ByVal X&, ByVal Y&, ByVal Flags&) As Long
    '
  
Sub iniciarListas(listview As listview, imagelist As imagelist, picture As PictureBox)
If iniciado = False Then
    Set fileTypes = New Collection
    Set myPic = picture
    Set myImageList = imagelist
    'myPic.ScaleMode = vbPixels
    myPic.ScaleWidth = 16
    myPic.ScaleHeight = 16
    myPic.picture = Form1.Icon
    myImageList.ImageHeight = 16
    myImageList.ImageWidth = 16
    SavePicture myPic, "@ME.ICO"
    myImageList.ListImages.Add , "@ME", myPic
    Set listview.SmallIcons = imagelist
    iniciado = True
End If

End Sub
Function loadFile(sFileName) As Integer
    Dim hImg    As Long     ' The handle to the system image list
    Dim r       As Long
    Dim sTemp   As String
    Dim iPos    As Integer
    Dim tInfo As SHFILEINFO
    
    myPic.picture = LoadPicture()    'clear it
    hImg = SHGetFileInfo(sFileName, 0&, tInfo, Len(tInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'mytype = tInfo.szTypeName
'iPos = InStr(mytype, vbNullChar)
'mytype = Mid$(mytype, 1, iPos - 1)
iicono = tInfo.iIcon
If noExiste(iicono) = True Then

fileIcon = iconExists(iicono)
If fileIcon <> "" Then
    myPic.picture = LoadPicture(fileIcon)
Else
    r& = ImageList_Draw(hImg, iicono, myPic.hDC, 0, 0, ILD_TRANSPARENT)
    With myPic               '1999/12/07 Now make them pictures, could be loaded into an image list. Larry Rebich
        .picture = .Image
    End With
    SavePicture myPic, iicono & ".ic_"
End If
    
    myImageList.ListImages.Add , "@" & iicono, myPic
    anadirTipo iicono
End If

loadFile = myImageList.ListImages.item("@" & iicono).Index
End Function


Function noExiste(cadena) As Boolean
For Each i In fileTypes
If i = cadena Then noExiste = False: Exit Function
Next
noExiste = True
End Function

Sub anadirTipo(cadena)
fileTypes.Add cadena, "@" & cadena
End Sub


Function iconExists(mytype)
If Dir$(App.path & "\" & mytype & ".ic_") <> "" Then
    iconExists = mytype & ".ic_"
Else
    iconExists = ""
End If
End Function

Sub clearIcons()
On Error Resume Next
For Each i In fileTypes
    Kill App.path & "\" & i & ".ic_"
Next
End Sub

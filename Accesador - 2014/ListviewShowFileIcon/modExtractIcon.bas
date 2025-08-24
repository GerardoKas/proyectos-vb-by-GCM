Attribute VB_Name = "modExtractIcon"
' modExtractIcon
' 2003/03/05 Copyright © 2003, Larry Rebich, using the DELL7500
' 2003/03/05 larry@larryrebich.com, www.larryrebich.com, 760-771-4730
' 2003/03/05 Made the April 2003 Tip of the Month

    Option Explicit
    DefLng A-Z
    
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
    
Public Sub LV_AddFileToListView(sPath As String, sFileName As String, sKeyPrefix As String, list As MSComctlLib.ListView, iTypeColumn As Integer, _
    imgLarge As MSComctlLib.ImageList, imgSmall As MSComctlLib.ImageList, picLarge As PictureBox, picSmall As PictureBox)
    
    '2003/11/30 Add sKeyPrefix - would fail if the filename is all numbers. Suggest use something like "@"
    
    Dim itm As ListItem
    Dim sType As String 'file type returned by API
    Dim sKey As String  '2003/11/30 Added
    
    SetSizePictureBoxes picLarge, picSmall  'need to be a certain size
    
    sKey = sKeyPrefix & sFileName
    
    With list
        .Icons = imgLarge
        .SmallIcons = imgSmall
        .ListItems.Add , sKey, sFileName    'add the file and use it's name as the key
        Set itm = .ListItems(sKey)          'point to the listitem
        AddImage itm, sPath & sFileName, sType, picLarge, picSmall, imgLarge, imgSmall
        If iTypeColumn > -1 Then            'display the file description in this column
            itm.SubItems(iTypeColumn) = sType
        End If
    End With
    
    Set itm = Nothing
End Sub
    
Public Sub LV_ClearImageListsAndRelinkToListView(list As MSComctlLib.ListView, imgLarge As ImageList, imgSmall As ImageList)
' 2003/03/06 Sub created by Larry Rebich while in La Quinta, CA.
    Dim img As ListImage
    Dim i As Integer
    
    Set list.Icons = Nothing                'unlink
    Set list.SmallIcons = Nothing
    
    ' remove all but the first one which is a dummy image
    For i = imgLarge.ListImages.Count - 1 To 2 Step -1
        imgLarge.ListImages.Remove i
    Next
    For i = imgSmall.ListImages.Count - 1 To 2 Step -1
        imgSmall.ListImages.Remove i
    Next
    
    list.Icons = imgLarge                   'relink
    list.SmallIcons = imgSmall

End Sub

Private Function ExtractIconIntoPictureUsingFile(sFileName As String, pic As PictureBox, bLarge As Boolean, tSHFILEINFO As SHFILEINFO) As Boolean
    Dim hImg    As Long     ' The handle to the system image list
    Dim r       As Long
    Dim sTemp   As String
    Dim iPos    As Integer
    
    On Error GoTo LV_ExtractIconIntoPictureUsingFileExit
    ' Set the pictureboxes to receive the icons.
    pic.Picture = LoadPicture()    'clear it
   
    ' Draw the associated icons into the picture boxes
    If bLarge Then
        hImg = SHGetFileInfo(sFileName, 0&, tSHFILEINFO, Len(tSHFILEINFO), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    Else
        hImg = SHGetFileInfo(sFileName, 0&, tSHFILEINFO, Len(tSHFILEINFO), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    End If
    r& = ImageList_Draw(hImg, tSHFILEINFO.iIcon, pic.hDC, 0, 0, ILD_TRANSPARENT)
    With pic               '1999/12/07 Now make them pictures, could be loaded into an image list. Larry Rebich
        .Picture = .Image
    End With
    
    With tSHFILEINFO
        iPos = InStr(.szTypeName, vbNullChar)
        If iPos > 0 Then
            .szTypeName = Mid$(.szTypeName, 1, iPos - 1)
        End If
    End With
    ExtractIconIntoPictureUsingFile = True
LV_ExtractIconIntoPictureUsingFileExit:
End Function

Private Function LV_InImageList(imgS As ListImages, sKey As String) As Boolean
' 2001/07/25 Function created by Larry Rebich while on Trip 2001 USA in Levis, Quebec
    Dim img As ListImage
    For Each img In imgS
        With img
            If .Key = sKey Then
                LV_InImageList = True
                Exit Function
            End If
        End With
    Next
End Function

Private Sub AddImage(itm As ListItem, sFileName As String, ByRef sType As String, picLarge As PictureBox, picSmall As PictureBox, _
    imgLarge As MSComctlLib.ImageList, imgSmall As MSComctlLib.ImageList)
    Dim tLarge As SHFILEINFO
    Dim tSmall As SHFILEINFO
    
    If ExtractIconIntoPictureUsingFile(sFileName, picLarge, True, tLarge) Then
        With tLarge
            If .iIcon > 0 Then
                If NotExistInImageList(imgLarge, .iIcon) Then
                    imgLarge.ListImages.Add , "K" & .iIcon, picLarge
                End If
                itm.Icon = "K" & .iIcon
            End If
        End With
    End If
    If ExtractIconIntoPictureUsingFile(sFileName, picSmall, False, tSmall) Then
        With tSmall
            If .iIcon > 0 Then
                If NotExistInImageList(imgSmall, .iIcon) Then
                    imgSmall.ListImages.Add , "K" & .iIcon, picSmall
                End If
                itm.SmallIcon = "K" & .iIcon
            End If
        End With
    End If
    sType = Trim$(tLarge.szTypeName)
End Sub

Private Function NotExistInImageList(imglst As ImageList, iIcon As Long) As Boolean
    Dim img As ListImage
    For Each img In imglst.ListImages
        With img
            If .Key = "K" & iIcon Then
                Exit Function
            End If
        End With
    Next
    NotExistInImageList = True
End Function

Private Sub SetSizePictureBoxes(picLarge As PictureBox, picSmall As PictureBox)
' 2003/03/06 Sub created by Larry Rebich while in La Quinta, CA.
    With picLarge               'make sure the picture boxes are the 'correct' size.
        .Move 0, 0, 480, 480
        .Visible = False
    End With
    With picSmall
        .Move 0, 0, 240, 240
        .Visible = False
    End With
End Sub

'Public Sub LV_AddImageAndPictureControls(frm As Form)
'    Dim picL As VB.PictureBox
'    Dim picS As VB.PictureBox
'    Dim imgL As MSComctlLib.ImageList
'    Dim imgS As MSComctlLib.ImageList
'    Set picL = frm.Controls.Add("VB.PictureBox", "picLarge", frm)
'    Set picS = frm.Controls.Add("VB.PictureBox", "picSmall", frm)
'    Set imgL = frm.Controls.Add("MSComctlLib.ImageListCtrl.2", "imgLarge", frm)
'    Set imgS = frm.Controls.Add("MSComctlLib.ImageListCtrl.2", "imgSmall", frm)
'
'    With imgL
'        .ImageHeight = 32
'        .ImageWidth = 32
'    End With
'    With imgS
'        .ImageHeight = 16
'        .ImageWidth = 16
'    End With
'End Sub







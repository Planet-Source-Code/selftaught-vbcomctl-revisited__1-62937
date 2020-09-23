Attribute VB_Name = "mResources"
'---------------------------------------------------------------------------------------
'mResources.bas           3/31/05
'
'            PURPOSE:
'               Extract bitmap resources from the .res file.
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Enum eBitmapResources
    bmpLarge256 = 101
    bmpSmall
    bmpPopupBack
    bmpPopupSide
    bmpRichEdit
    bmpHelpToolbar
End Enum

Private Function pTrue(ByRef b As Boolean) As Boolean
    b = True
    pTrue = True
End Function

Public Property Get InIDE() As Boolean
    Debug.Assert pTrue(InIDE)
End Property

Private Sub pAddImage(ByVal oIml As cImageList, ByVal iId As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Add an image to the imagelist directly from the resource if
'             compiled, or from a StdPicture object if in the ide.
'---------------------------------------------------------------------------------------
    If InIDE() _
        Then oIml.AddFromHandle LoadResPicture(iId, vbResBitmap).Handle, imlBitmap _
        Else oIml.AddFromResource iId, imlBitmap, App.hInstance
End Sub

Public Function gImageListRichEdit() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the imagelist used by the richedit toolbar.
'---------------------------------------------------------------------------------------
    Static oImageList As cImageList
    If oImageList Is Nothing Then
        Set oImageList = NewImageList(16, 16, imlColor8)
        pAddImage oImageList, bmpRichEdit
    End If
    Set gImageListRichEdit = oImageList
End Function

Public Function gImageListLarge() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the general 24x24 imagelist.
'---------------------------------------------------------------------------------------
    Static oImageList As cImageList
    If oImageList Is Nothing Then
        Set oImageList = NewImageList(24, 24, imlColor8)
        pAddImage oImageList, bmpLarge256
    End If
    Set gImageListLarge = oImageList
End Function

Public Function gImageListSmall() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the general 16x16 imagelist.
'---------------------------------------------------------------------------------------
    Static oImageList As cImageList
    If oImageList Is Nothing Then
        Set oImageList = NewImageList(16, 16, imlColor8)
        pAddImage oImageList, bmpSmall
    End If
    Set gImageListSmall = oImageList
End Function

Public Function gImageListHelp() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the imagelist with the help contents, search and
'             usercontrol icons.
'---------------------------------------------------------------------------------------
    Static oImageList As cImageList
    If oImageList Is Nothing Then
        Set oImageList = NewImageList(16, 16, imlColor8)
        pAddImage oImageList, bmpHelpToolbar
    End If
    Set gImageListHelp = oImageList
End Function

Public Function gSysImageListSmall() As cSysImageList
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the 16x16 system imagelist.
'---------------------------------------------------------------------------------------
    Static oImageList As cSysImageList
    If oImageList Is Nothing Then
        Set oImageList = NewSysImageList()
        pInitSystemImagelist oImageList
    End If
    Set gSysImageListSmall = oImageList
End Function

Public Function gSysImageListLarge() As cSysImageList
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the 32x32 system imagelist.
'---------------------------------------------------------------------------------------
    Static oImageList As cSysImageList
    If oImageList Is Nothing Then
        Set oImageList = NewSysImageList(True)
        pInitSystemImagelist oImageList
    End If
    Set gSysImageListLarge = oImageList
End Function

Private Sub pInitSystemImagelist(ByVal oIml As cSysImageList)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Make sure that the new system imagelist has at least a few icons.
'---------------------------------------------------------------------------------------
    With oIml
        .ItemIndex "*.zip", True
        .ItemIndex "*.txt", True
        .ItemIndex "*.pdf", True
        .ItemIndex "*.rtf", True
        .ItemIndex "*.doc", True
    End With
End Sub

Public Property Get PopupBackPicture() As StdPicture
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the picture used as the background picture for the popup menus.
'---------------------------------------------------------------------------------------
    Set PopupBackPicture = LoadResPicture(bmpPopupBack, vbResBitmap)
End Property

Public Property Get PopupSideBar() As StdPicture
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the picture used as a sidebar for the popup menus.
'---------------------------------------------------------------------------------------
    Set PopupSideBar = LoadResPicture(bmpPopupSide, vbResBitmap)
End Property

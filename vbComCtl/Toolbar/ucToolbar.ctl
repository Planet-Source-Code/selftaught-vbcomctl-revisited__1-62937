VERSION 5.00
Begin VB.UserControl ucToolbar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   HasDC           =   0   'False
   PropertyPages   =   "ucToolbar.ctx":0000
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   ToolboxBitmap   =   "ucToolbar.ctx":000D
End
Attribute VB_Name = "ucToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucToolbar.ctl        12/15/04
'
'           PURPOSE:
'               Implement the comctl32 toolbar control.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Toolbar/vbAccelerator_ToolBar_and_CoolMenu_Control/VB6_Toolbar_Complete_Source.asp
'               cToolbar.ctl
'
'==================================================================================================

Option Explicit

Implements iSubclass

Public Enum eToolbarButtonStyle
    tbarButtonNormal = BTNS_BUTTON
    tbarButtonSeparator = BTNS_SEP
    tbarButtonCheck = BTNS_CHECK
    tbarButtonCheckGroup = BTNS_CHECKGROUP
    tbarButtonDropDown = BTNS_DROPDOWN
    tbarButtonWholeDropDown = BTNS_WHOLEDROPDOWN
End Enum

Public Enum eToolbarImageSource
    tbarDefImageCustom = NegOneL
    tbarDefImageStandardSmall = IDB_STD_SMALL_COLOR
    tbarDefImageStandardLarge = IDB_STD_LARGE_COLOR
    tbarDefImageViewSmall = IDB_VIEW_SMALL_COLOR
    tbarDefImageViewLarge = IDB_VIEW_LARGE_COLOR
    tbarDefImageHistorySmall = IDB_HIST_SMALL_COLOR
    tbarDefImageHistoryLarge = IDB_HIST_LARGE_COLOR
End Enum

Public Enum eToolbarStandardImage
    tbarStdImageCut = STD_CUT
    tbarStdImageCopy = STD_COPY
    tbarStdImagePaste = STD_PASTE
    tbarStdImageUndo = STD_UNDO
    tbarStdImageRedo = STD_REDOW
    tbarStdImageDelete = STD_DELETE
    tbarStdImageFileNew = STD_FILENEW
    tbarStdImageFileOpen = STD_FILEOPEN
    tbarStdImageFileSave = STD_FILESAVE
    tbarStdImagePrintPreview = STD_PRINTPRE
    tbarStdImageProperties = STD_PROPERTIES
    tbarStdImageHelp = STD_HELP
    tbarStdImageFind = STD_FIND
    tbarStdImageReplace = STD_REPLACE
    tbarStdImagePrint = STD_PRINT
End Enum

Public Enum eToolbarViewImage
    tbarViewImageLargeIcons = VIEW_LARGEICONS
    tbarViewImageSmallIcons = VIEW_SMALLICONS
    tbarViewImageList = VIEW_LIST
    tbarViewImageDetails = VIEW_DETAILS
    tbarViewImageSortName = VIEW_SORTNAME
    tbarViewImageSortSize = VIEW_SORTSIZE
    tbarViewImageSortDate = VIEW_SORTDATE
    tbarViewImageSortType = VIEW_SORTTYPE
    tbarViewImageParentFolder = VIEW_PARENTFOLDER
    tbarViewImageNetConnect = VIEW_NETCONNECT
    tbarViewImageNetDisconnect = VIEW_NETDISCONNECT
    tbarViewImageNewFolder = VIEW_NEWFOLDER
'#if (_WIN32_IE >= 0x0400)
    tbarViewImageMenu = VIEW_VIEWMENU
'#End If
End Enum

Public Enum eToolbarPopupPosition
    tbarPopLiteral = vbAlignNone
    tbarPopRightDown = vbAlignTop
    tbarPopRightUp = vbAlignBottom
    tbarPopBottomRight = vbAlignLeft
    tbarPopBottomLeft = vbAlignRight
End Enum

Public Enum eToolbarHistoryImage
    tbarHistoryImageBack = HIST_BACK
    tbarHistoryImageForward = HIST_FORWARD
    tbarHistoryImageFavorites = HIST_FAVORITES
    tbarHistoryImageAddToFavorites = HIST_ADDTOFAVORITES
    tbarHistoryImageViewTree = HIST_VIEWTREE
End Enum

Public Event ButtonClick(ByVal oButton As cButton)
Public Event ButtonDropDown(ByVal oButton As cButton)
Public Event RightButtonDown(ByVal x As Single, ByVal y As Single)
Public Event RightButtonUp(ByVal x As Single, ByVal y As Single)
Public Event Resize()
Public Event ExitMenuTrack()

Private Const UM_TRACKMENU As Long = WM_USER + &H66BC&
Private Const UM_SHOWCHEVRON As Long = WM_USER + &H66BF&

Private Type tToolbarButton
    iId                 As Long
    sKey                As String
    sToolTip            As String
End Type

Private Const PROP_Font = "Font"
Private Const PROP_Style = "Style"
Private Const PROP_Vertical = "Vertical"
Private Const PROP_ButtonWidth = "ButtonWidth"
Private Const PROP_ImageSource = "ImageSource"
Private Const PROP_TextRows = "TextRows"
Private Const PROP_MenuStyle = "MenuStyle"
Private Const PROP_Themeable = "Themeable"

Private Const DEF_Style = TBSTYLE_FLAT
Private Const DEF_Vertical = False
Private Const DEF_ButtonWidth As Long = &H800010
Private Const DEF_ImageSource As Long = tbarDefImageCustom
Private Const DEF_TextRows As Long = 1
Private Const DEF_MenuStyle As Boolean = False
Private Const DEF_Themeable = True

Private WithEvents moFont As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private WithEvents moImageListEvent As cImageList
Attribute moImageListEvent.VB_VarHelpID = -1
Private moImageList     As cImageList

Private mtButton        As TBBUTTON
Private mtButtonInfo    As TBBUTTONINFO

Private mtButtons()     As tToolbarButton
Private miButtonCount   As Long
Private miButtonUbound  As Long
Private miButtonControl As Long

Private miButtonWidth   As Long

Private mhWnd           As Long
Private moTrackMenu     As pcTrackToolMenu
Private mhFont          As Long
Private miImageSource   As eToolbarImageSource
Private miTextRows      As Long

Private mbVertical      As Boolean
Private mbThemeable     As Boolean

Private miStyle         As Long

Private msTextBuffer    As String * 130

Private Const ucToolbar = "ucToolbar"
Private Const cButton = "cButton"
Private Const cButtons = "cButtons"
 
Private Const NMHDR_hwndFrom    As Long = 0
'Private Const NMHDR_idfrom      As Long = 4
Private Const NMHDR_code        As Long = 8
 
Private Const NMTOOLBAR_iItem      As Long = 12
'Private Const NMTOOLBAR_lpszString As Long = 40
'Private Const NMTOOLBAR_cchText    As Long = 36
 
Private Const NMTBGETINFOTIP_pszText    As Long = 12
Private Const NMTBGETINFOTIP_cchText    As Long = 16
Private Const NMTBGETINFOTIP_iItem      As Long = 20

Private Const NMTBHOTITEM_iNew          As Long = 16

Private moRebar As ucRebar
Private miRebarId As Long
Private mbRebarVertical As Boolean

Private moChevron As pcToolbarChevron
Private mhWndChevron As Long
Private moDroppedMenuButton As cButton
Private mbIgnoreMenuKey As Boolean

Private miIgnoreExitTrack As Long
Private mbTrackFromDropDown As Boolean

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Respond to notifications from the toolbar.
'---------------------------------------------------------------------------------------
    Dim liIndex As Long
    
    Select Case uMsg
    Case WM_NOTIFY
        bHandled = True
        lReturn = ZeroL
        Select Case MemOffset32(lParam, NMHDR_code)
        Case TBN_GETINFOTIPA
            liIndex = pCommandToIndex(MemOffset32(lParam, NMTBGETINFOTIP_iItem))
            If liIndex > NegOneL Then
                With mtButtons(liIndex)
                    liIndex = MemOffset32(lParam, NMTBGETINFOTIP_cchText)
                    If liIndex > ZeroL Then
                        If LenB(.sToolTip) < liIndex Then liIndex = LenB(.sToolTip)
                        If liIndex > ZeroL Then
                            CopyMemory ByVal MemOffset32(lParam, NMTBGETINFOTIP_pszText), ByVal StrPtr(.sToolTip), liIndex
                            If MemOffset32(lParam, NMTBGETINFOTIP_cchText) > liIndex Then liIndex = liIndex + OneL
                            MemOffset16(MemOffset32(lParam, NMTBGETINFOTIP_pszText), liIndex) = 0
                        End If
                    End If
                End With
            End If
        Case TBN_DROPDOWN
            liIndex = pCommandToIndex(MemOffset32(lParam, NMTOOLBAR_iItem))
            If liIndex > NegOneL Then
                If Not moTrackMenu Is Nothing Then
                    PostMessage UserControl.hWnd, UM_TRACKMENU, tbarTrackDropped Or &H10000, liIndex
                Else
                    moChevron.fSuspend True
                    RaiseEvent ButtonDropDown(pButton(liIndex))
                    moChevron.fSuspend False
                End If
            End If
        Case TBN_HOTITEMCHANGE
            If Not moTrackMenu Is Nothing Then
                liIndex = pCommandToIndex(MemOffset32(lParam, NMTBHOTITEM_iNew))
                If mhWndChevron Then
                    If MemOffset32(lParam, NMHDR_hwndFrom) = mhWndChevron Then
                        If moTrackMenu.fTrackIndex >= (SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) - SendMessage(mhWndChevron, TB_BUTTONCOUNT, ZeroL, ZeroL)) Then
                            If moTrackMenu.fTrackState = tbarTrackButtons Then moTrackMenu.fHotItemChange liIndex, lReturn
                        End If
                    Else
                        If moTrackMenu.fTrackState = tbarTrackButtons Then moTrackMenu.fHotItemChange liIndex, lReturn
                    End If
                Else
                    If moTrackMenu.fTrackState = tbarTrackButtons Then moTrackMenu.fHotItemChange liIndex, lReturn
                End If
            End If
        End Select
        
    Case UM_TRACKMENU
        bHandled = True
        moChevron.fSuspend True
        mbTrackFromDropDown = CBool(wParam And &HFFFF0000)
        moTrackMenu.fTrack wParam And &HFFFF&, lParam
        mbTrackFromDropDown = False
        moChevron.fSuspend False
        
    Case UM_SHOWCHEVRON
        bHandled = True
        miIgnoreExitTrack = miIgnoreExitTrack - OneL
        PostMessage UserControl.hWnd, UM_TRACKMENU, wParam, lParam
        If Not moRebar Is Nothing Then moRebar.fToolbar_ShowChevron Me, miRebarId
        
    Case WM_TIMER
        moTrackMenu.fTimer
        
    Case WM_COMMAND
        liIndex = pCommandToIndex(wParam)
        If liIndex > NegOneL Then
            If mhWndChevron Then
                If lParam = mhWndChevron Then
                    pOnChevronClick wParam
                End If
            End If
            
            'RDown-LDown-Lup-RUp causes this on at least CC 5.80 and CC 6.
            'Click events should not be raised for whole drop down buttons.
            If pButton_Info(wParam, TBIF_STYLE) <> tbarButtonWholeDropDown Then
                RaiseEvent ButtonClick(pButton(liIndex))
            End If
        End If
        
    Case WM_ERASEBKGND
        bHandled = True
    
    Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
        bHandled = pOnLButtonDown(hWnd, lParam)
    
    Case WM_RBUTTONDOWN
        bHandled = True
        If hWnd = mhWnd _
            Then RaiseEvent RightButtonDown(loword(lParam), hiword(lParam))
        
    Case WM_RBUTTONUP
        bHandled = True
        If hWnd = mhWnd _
            Then pRButtonUp loword(lParam), hiword(lParam)
            
    Case WM_SIZE
        pResize
        
    End Select
End Sub

Private Sub moImageListEvent_Changed()
    pSetImageList mhWnd
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
    LoadShellMod
    InitCC ICC_BAR_CLASSES
    mtButtonInfo.cbSize = LenB(mtButtonInfo)
    Set moChevron = New pcToolbarChevron
    Set moFontPage = New pcSupportFontPropPage
    ForceWindowToShowAllUIStates UserControl.hWnd
End Sub

Private Sub UserControl_InitProperties()
    Set moFont = Font_CreateDefault(Ambient.Font)
    miStyle = DEF_Style
    mbVertical = DEF_Vertical
    miButtonWidth = DEF_ButtonWidth
    miImageSource = DEF_ImageSource
    miTextRows = DEF_TextRows
    #If DEF_MenuStyle Then
        Set moTrackMenu = New pcTrackToolMenu
        Set moTrackMenu.fOwner = Me
    #End If
    mbThemeable = DEF_Themeable
    pCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    miStyle = PropBag.ReadProperty(PROP_Style, DEF_Style)
    mbVertical = PropBag.ReadProperty(PROP_Vertical, DEF_Vertical)
    miButtonWidth = PropBag.ReadProperty(PROP_ButtonWidth, DEF_ButtonWidth)
    miImageSource = PropBag.ReadProperty(PROP_ImageSource, DEF_ImageSource)
    miTextRows = PropBag.ReadProperty(PROP_TextRows, DEF_TextRows)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    
    If PropBag.ReadProperty(PROP_MenuStyle, DEF_MenuStyle) Then
        Set moTrackMenu = New pcTrackToolMenu
        Set moTrackMenu.fOwner = Me
    End If
    
    pCreate
    pSetImageSource
End Sub

Private Sub UserControl_Terminate()
    TrackKey.StopTrack moTrackMenu
    Set moTrackMenu = Nothing
    Set moFontPage = Nothing
    Set moChevron = Nothing
    pDestroy
    ReleaseShellMod
    If mhFont Then moFont.ReleaseHandle mhFont
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_Style, miStyle, DEF_Style
    PropBag.WriteProperty PROP_Vertical, mbVertical, DEF_Vertical
    PropBag.WriteProperty PROP_ButtonWidth, miButtonWidth, DEF_ButtonWidth
    PropBag.WriteProperty PROP_ImageSource, miImageSource, DEF_ImageSource
    PropBag.WriteProperty PROP_TextRows, miTextRows, DEF_TextRows
    PropBag.WriteProperty PROP_MenuStyle, Not moTrackMenu Is Nothing, DEF_MenuStyle
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
End Sub

Private Sub moFont_Changed()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Update the font used by the toolbar.
'---------------------------------------------------------------------------------------
    moFont.OnAmbientFontChanged Ambient.Font
    pSetFont
    pPropChanged PROP_Font
End Sub

Private Function pOnLButtonDown(ByVal lhWnd As Long, ByVal lParam As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return whether the lbuttondown event is over a checked item in a button group.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim tP As POINT
        tP.x = lParam And &HFFFF&
        tP.y = ((lParam And &HFFFF0000) \ &H10000) And &HFFFF&
        Dim liIndex As Long
        Dim liId As Long
        liIndex = SendMessageAny(lhWnd, TB_HITTEST, ZeroL, tP)
        
        If liIndex > ZeroL Then
            If SendMessageAny(lhWnd, TB_GETBUTTON, liIndex, mtButton) Then
                If mtButton.fsStyle = BTNS_CHECKGROUP Then
                    pOnLButtonDown = CBool(CLng(mtButton.fsState) And TBSTATE_CHECKED)
                End If
            End If
        End If
    End If
End Function

Private Sub pRButtonUp(ByVal x As Long, ByVal y As Long)
    Dim tP As POINT
    tP.x = x
    tP.y = y
    MapWindowPoints mhWnd, ZeroL, tP, OneL
    If WindowFromPoint(tP.x, tP.y) = mhWnd Then
        RaiseEvent RightButtonUp(x, y)
    End If
    
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    Set fSupportFontPropPage = moFontPage
End Property

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
    Set o = Ambient.Font
End Sub

Private Sub pOnChevronClick(ByVal iId As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Sync the toolbar check state with the chevron if necessary.
'---------------------------------------------------------------------------------------
    Dim liState As Long
    
    liState = pButton_Info(iId, TBIF_STATE)
    
    With mtButtonInfo

        .dwMask = TBIF_STATE
        SendMessage mhWndChevron, TB_GETBUTTONINFO, iId, VarPtr(mtButtonInfo)
        
        If (liState And TBSTATE_CHECKED) _
            Xor (.fsState And TBSTATE_CHECKED) _
                Then SendMessage mhWnd, TB_CHECKBUTTON, iId, -CBool(.fsState And TBSTATE_CHECKED)

    End With
End Sub

Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Create the toolbar and install the subclasses.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode Then
        pDestroy
        
        pSetFont
        
        mhWnd = pCreateToolbar()
        
        If mhWnd Then
        
            pResize
            
            ShowWindow mhWnd, SW_SHOWNORMAL
            
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_COMMAND, WM_ERASEBKGND, UM_TRACKMENU, UM_SHOWCHEVRON, WM_TIMER, WM_SIZE)
            pSubclassToolbar mhWnd, True
        End If
    End If
End Sub

Private Sub pSubclassToolbar(ByVal lhWnd As Long, ByVal bVal As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Subclass a toolbar if we're on commctrl version 6.
'---------------------------------------------------------------------------------------
    If lhWnd Then
        If bVal _
            Then Subclass_Install Me, lhWnd, Array(WM_RBUTTONDOWN, WM_RBUTTONUP, WM_LBUTTONDOWN, WM_LBUTTONDBLCLK) _
            Else Subclass_Remove Me, lhWnd
    End If
End Sub

Private Sub pRecreate(ByVal bVertical As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Recreate the toolbar with the same buttons and style.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode Then
        Dim lhWnd As Long
        lhWnd = pCreateToolbar()
        
        If lhWnd Then
        
            Dim liStyle As Long
            liStyle = GetWindowLong(lhWnd, GWL_STYLE)
            
            liStyle = liStyle And Not (CCS_TOP Or CCS_RIGHT)
            If bVertical Then
                liStyle = liStyle Or CCS_LEFT
            Else
                liStyle = liStyle Or CCS_TOP
            End If
            SetWindowLong lhWnd, GWL_STYLE, liStyle
            SendMessage lhWnd, TB_SETSTYLE, ZeroL, liStyle
            
            pCopyButtons mhWnd, lhWnd, , , (bVertical + OneL) * TBSTATE_WRAP
            
            pSubclassToolbar mhWnd, False
            DestroyWindow mhWnd
            
            mhWnd = lhWnd
            pSubclassToolbar mhWnd, True
    
            'pSetFont
            pResize
            
            ShowWindow mhWnd, SW_SHOWNORMAL
            
        End If
    End If
End Sub

Private Function pFirstClippedIndex()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the first index in the toolbar which is not visible.
'---------------------------------------------------------------------------------------
    
    If mhWnd Then
        
        Dim ltClient As RECT
        GetClientRect mhWnd, ltClient
        ltClient.Right = ltClient.Right + OneL
        ltClient.bottom = ltClient.bottom + OneL
        
        Dim ltButton As RECT
        
        For pFirstClippedIndex = SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) - OneL To ZeroL Step NegOneL
            If Not pIsSep(pFirstClippedIndex) Then
                If SendMessage(mhWnd, TB_GETITEMRECT, pFirstClippedIndex, VarPtr(ltButton)) Then
                    If PtInRect(ltClient, ltButton.Right, ltButton.bottom) Then Exit For
                End If
            End If
        Next
        
        pFirstClippedIndex = pFirstClippedIndex + OneL
        
    End If
End Function

Private Function pCreateToolbar() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Create a toolbar with the correct style and initialization.
'---------------------------------------------------------------------------------------
    Dim lsAnsi As String
    lsAnsi = StrConv(WC_TOOLBAR & vbNullChar, vbFromUnicode)
    pCreateToolbar = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, pRealWinStyle And Not WS_VISIBLE, ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    Debug.Assert pCreateToolbar
    
    If pCreateToolbar Then
        SendMessage pCreateToolbar, TB_BUTTONSTRUCTSIZE, LenB(mtButton), ZeroL
        SendMessage pCreateToolbar, TB_SETEXTENDEDSTYLE, ZeroL, -CBool(moTrackMenu Is Nothing) * TBSTYLE_EX_DRAWDDARROWS
        'SendMessage pCreateToolbar, TB_SETBUTTONWIDTH, ZeroL, miButtonWidth
        SendMessage pCreateToolbar, TB_SETMAXTEXTROWS, miTextRows, ZeroL
        
        pSetImageList pCreateToolbar
        If mhFont Then SendMessage pCreateToolbar, WM_SETFONT, mhFont, OneL
        
        EnableWindowTheme pCreateToolbar, mbThemeable
    End If
End Function

Private Sub pCopyButtons(ByVal hWndFrom As Long, ByVal hWndTo As Long, Optional ByVal iFirstButton As Long = NegOneL, Optional ByVal iStateOr As Long, Optional ByVal iStateAndNot As Long, Optional ByVal iStyleOr As Long, Optional ByVal iStyleAndNot As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Copy buttons from one toolbar to another.
'---------------------------------------------------------------------------------------
    
    'if iFirstButton < 0 then it is assumed we are creating a chevron menu
    Dim i As Long
    If iFirstButton > NegOneL Then
        
        Do While pIsSep(iFirstButton)
            iFirstButton = iFirstButton + OneL
        Loop
        
    End If
    
    For i = IIf(iFirstButton <= ZeroL, ZeroL, iFirstButton) To SendMessage(hWndFrom, TB_BUTTONCOUNT, ZeroL, ZeroL) - OneL
        If SendMessage(hWndFrom, TB_GETBUTTON, i, VarPtr(mtButton)) Then
            With mtButton
                .fsStyle = ((.fsStyle Or iStyleOr) And Not iStyleAndNot) And Not -CBool(.fsStyle = tbarButtonSeparator) * BTNS_AUTOSIZE
                .fsState = (.fsState Or iStateOr) And Not iStateAndNot
                If CBool(.fsStyle And tbarButtonSeparator) Then
                    .fsStyle = .fsStyle And Not BTNS_AUTOSIZE
                End If
            End With
            SendMessage hWndTo, TB_ADDBUTTONSW, OneL, VarPtr(mtButton)
        Else
            Debug.Assert False
        End If
    Next
    
End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Destroy the toolbar and subclasses.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Subclass_Remove Me, UserControl.hWnd
        pSubclassToolbar mhWnd, False
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
End Sub

Private Function pCommandToIndex(ByVal iCmd As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Find the index of a button given its command id.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If iCmd Then
            pCommandToIndex = SendMessage(mhWnd, TB_COMMANDTOINDEX, iCmd, ZeroL)
            'Debug.Assert pCommandToIndex > NegOneL And pCommandToIndex < miButtonCount
            'Debug.Assert mtButtons(pCommandToIndex).iId = iCmd
        Else
            pCommandToIndex = NegOneL
        End If
    End If
End Function
Private Sub pResize()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Account for changes in alignment and resize the usercontrol.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode And moRebar Is Nothing Then
        If mhWnd Then
            
            Static b As Boolean
            If b Then Exit Sub
            
            pCheckVert
            
            On Error GoTo handler
            b = True
            SendMessage mhWnd, TB_AUTOSIZE, ZeroL, ZeroL
            Dim liWidth As Long
            Dim liHeight As Long
            Select Case pAlignment
            Case vbAlignTop, vbAlignBottom
                pGetIdealSize liWidth, liHeight, False
                Height = ScaleY(liHeight, vbPixels, vbTwips)
            Case vbAlignLeft, vbAlignRight
                pGetIdealSize liWidth, liHeight, True
                Width = ScaleX(liWidth, vbPixels, vbTwips)
            Case Else
                pGetIdealSize liWidth, liHeight, mbVertical
                SIZE ScaleX(liWidth, vbPixels, vbTwips), ScaleY(liHeight, vbPixels, vbTwips)
            End Select
            SendMessage mhWnd, TB_AUTOSIZE, ZeroL, ZeroL
            'If CheckCCVersion(6&) Then moWnd.Update
            b = False
            RaiseEvent Resize
            Exit Sub
handler:
            Debug.Print "Toolbar Resize Error: ", Err.Description
            Debug.Assert False
            Resume Next
        End If
    End If
End Sub

Private Sub pGetIdealSize(ByRef iWidth As Long, ByRef iHeight As Long, ByVal bVertical As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the ideal width and height of the toolbar to accomodate its buttons.
'---------------------------------------------------------------------------------------
    
    iWidth = ZeroL
    iHeight = ZeroL
    
'    Dim tS As tSize
'    If sendmessage(mhwnd, TB_GETMAXSIZE, ZeroL, VarPtr(tS)) Then
'        iWidth = tS.cx
'        iHeight = tS.cy
'    End If
    
    Dim tR As RECT
    Dim i As Long
    
    For i = ZeroL To SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) - OneL
        If Not pIsSep(i) Then
            If SendMessage(mhWnd, TB_GETITEMRECT, i, VarPtr(tR)) Then
                'If bVertical Then
                    If tR.Right > iWidth Then iWidth = tR.Right
                    If tR.bottom > iHeight Then iHeight = tR.bottom
                'End If
            End If
        End If
    Next
    
End Sub


Private Sub pCheckVert()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Ensure that the CCS_VERT style is correct and recreate the toolbar when
'             changing between vertical modes.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim lbVert As Boolean
        
        If moRebar Is Nothing Then
            Select Case pAlignment
            Case vbAlignNone
                lbVert = mbVertical
            Case vbAlignLeft, vbAlignRight
                lbVert = True
            End Select
            If CBool(GetWindowLong(mhWnd, GWL_STYLE) And CCS_VERT) Xor lbVert Then pRecreate lbVert
        End If
    End If
End Sub


Private Function pIsSep(ByVal iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the given button index is a separator.
'---------------------------------------------------------------------------------------
    If iIndex > NegOneL And iIndex < miButtonCount Then
        pIsSep = CBool(pButton_Info(mtButtons(iIndex).iId, TBIF_STYLE) And BTNS_SEP)
    End If
End Function

Private Function pAlignment() As evbComCtlAlignment
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the alignment mode of the usercontrol.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    Dim loExtender As VBControlExtender
    Set loExtender = Extender
    'early bound!
    pAlignment = loExtender.Align
    On Error GoTo 0
    Exit Function
handler:
    On Error Resume Next
    'Maybe it's still supported.
    'even if it is, it could have different return values. Oh well.
    pAlignment = Extender.Align
    On Error GoTo 0
End Function

Private Sub pSetImageList(ByVal lhWnd As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the imagelist or default images to a toolbar.
'---------------------------------------------------------------------------------------
    If lhWnd Then
        If Not moImageList Is Nothing Then
            SendMessage lhWnd, TB_SETIMAGELIST, ZeroL, moImageList.hIml
        Else
            If miImageSource = tbarDefImageCustom Then SendMessage lhWnd, TB_SETIMAGELIST, ZeroL, ZeroL
        End If
    End If
End Sub

Private Sub pSetFont()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the font of the usercontrol and resize the control.
'---------------------------------------------------------------------------------------
    'pResize
    Dim hFont As Long
    hFont = moFont.GetHandle()
    If hFont Then
        If mhWnd Then SendMessage mhWnd, WM_SETFONT, hFont, OneL
        If mhFont Then moFont.ReleaseHandle mhFont
        mhFont = hFont
        pResize
        If Not moRebar Is Nothing Then moRebar.fToolbar_Resize Me, miRebarId
    End If
End Sub

Private Sub pSetStyle(Optional ByVal iStyleOr As Long, Optional ByVal iStyleAndNot As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the style of the toolbar.
'---------------------------------------------------------------------------------------
    miStyle = (miStyle And Not iStyleAndNot) Or iStyleOr
    
    If mhWnd Then
        Dim liStyle As Long
        liStyle = pRealWinStyle()
        
        SetWindowLong mhWnd, GWL_STYLE, liStyle
        SendMessage mhWnd, TB_SETSTYLE, ZeroL, liStyle
        
    End If
End Sub

Private Function pRealWinStyle() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the window style for the control.
'---------------------------------------------------------------------------------------
    pRealWinStyle = WS_CHILD Or WS_VISIBLE Or TBSTYLE_TOOLTIPS Or CCS_NODIVIDER Or miStyle
    If Not moRebar Is Nothing Then
        pRealWinStyle = ((pRealWinStyle Or CCS_NOPARENTALIGN Or CCS_NORESIZE Or TBSTYLE_TRANSPARENT) And Not (CCS_TOP Or CCS_LEFT)) Or IIf(mbRebarVertical, CCS_LEFT, CCS_TOP)
        If mbRebarVertical Then pRealWinStyle = pRealWinStyle Or TBSTYLE_WRAPABLE Else pRealWinStyle = pRealWinStyle And Not TBSTYLE_WRAPABLE
    Else
        Select Case pAlignment
        Case vbAlignNone
            If mbVertical _
                Then pRealWinStyle = pRealWinStyle Or CCS_LEFT Or TBSTYLE_WRAPABLE _
                Else pRealWinStyle = pRealWinStyle Or CCS_TOP
        Case Is > vbAlignBottom
            pRealWinStyle = pRealWinStyle Or CCS_LEFT Or TBSTYLE_WRAPABLE
        Case Else
            pRealWinStyle = pRealWinStyle Or CCS_TOP
        End Select
    End If
End Function

Private Sub pPropChanged(ByRef s As String)
    If Ambient.UserMode = False Then PropertyChanged s
End Sub

Private Function pMapAccelerator(ByVal iChar As Integer, ByVal iStart As Long, ByRef bDuplicate As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Find a button with the specified char as its mnemonic, or a button whose
'             begins with the given char.
'---------------------------------------------------------------------------------------
                                        'vbKeyUp = asc("&")
    If CBool(mhWnd) And iChar <> vbKeyUp Then

'        Doesn't work...
'        If sendmessage(mhwnd, TB_MAPACCELERATORW, iChar, VarPtr(fMenu_MapAccelerator)) = ZeroL Then
'            fMenu_MapAccelerator = NegOneL
'        Else
'            fMenu_MapAccelerator = pCommandToIndex(fMenu_MapAccelerator)
'        End If
        
        If iStart < ZeroL Then iStart = ZeroL
        
        Dim i As Long
        Dim liPos As Long
        Dim lsText As String
        
        Dim liAccel() As Long
        Dim liAccelCount As Long
        Dim liLetter() As Long
        Dim liLetterCount As Long
        
        For i = ZeroL To miButtonCount - OneL
            lsText = pButton_Text(mtButtons(i).iId)
            
            If Asc(lsText) = iChar Then
                ReDim Preserve liLetter(0 To liLetterCount)
                liLetter(liLetterCount) = i
                liLetterCount = liLetterCount + OneL
            End If
            
            If AccelChar(lsText) = iChar Then
                ReDim Preserve liAccel(0 To liAccelCount)
                liAccel(liAccelCount) = i
                liAccelCount = liAccelCount + OneL
            End If
        Next
        
        pMapAccelerator = NegOneL
        
        If liAccelCount > ZeroL Then
            If liAccelCount > OneL Then
                For i = liAccelCount - OneL To ZeroL Step NegOneL
                    If liAccel(i) <= iStart Then Exit For
                    pMapAccelerator = liAccel(i)
                Next
                If pMapAccelerator = NegOneL Then pMapAccelerator = liAccel(0)
                bDuplicate = True
            Else
                pMapAccelerator = liAccel(0)
            End If
        ElseIf liLetterCount > ZeroL Then
            If liLetterCount > OneL Then
                For i = liLetterCount - OneL To ZeroL Step NegOneL
                    If liLetter(i) <= iStart Then Exit For
                    pMapAccelerator = liLetter(i)
                Next
                If pMapAccelerator = NegOneL Then pMapAccelerator = liLetter(0)
                bDuplicate = True
            Else
                pMapAccelerator = liLetter(0)
            End If
        End If
        
    End If
    
End Function

Private Sub pDeleteButton(ByVal iIndex As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove the button from the array.
'---------------------------------------------------------------------------------------
    Debug.Assert iIndex > NegOneL And iIndex < miButtonCount
    If iIndex > NegOneL And iIndex < miButtonCount Then
        miButtonCount = miButtonCount - OneL
        If iIndex < miButtonCount Then
            With mtButtons(iIndex)
                .sKey = vbNullString
                .sToolTip = vbNullString
                CopyMemory .iId, mtButtons(iIndex + OneL).iId, LenB(mtButtons(0)) * (miButtonCount - iIndex)
            End With
            ZeroMemory mtButtons(miButtonCount).iId, LenB(mtButtons(0))
        End If
        pResize
    End If
End Sub

Private Function pButton(ByVal iIndex As Long) As cButton
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a cButton object identifying the given index.
'---------------------------------------------------------------------------------------
    'Debug.Assert iIndex > NegOneL And iIndex < miButtonCount
    
    If iIndex > NegOneL And iIndex < miButtonCount Then
        Set pButton = New cButton
        pButton.fInit Me, mtButtons(iIndex).iId, iIndex
    End If
End Function


Friend Sub fButtons_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the next cButton object in an enumeration.
'---------------------------------------------------------------------------------------
    With tEnum
        .iIndex = .iIndex + OneL
        If .iIndex > NegOneL And .iIndex < miButtonCount Then
            Set vNextItem = pButton(.iIndex)
        Else
            bNoMoreItems = True
        End If
    End With
End Sub

Friend Property Get fButtons_Control() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return an identity number for the current buttons collection.
'---------------------------------------------------------------------------------------
    fButtons_Control = miButtonControl
End Property

Friend Function fButtons_Add( _
            ByRef sKey As String, _
            ByRef sText As String, _
            ByVal iStyle As eToolbarButtonStyle, _
            ByVal iIconIndex As Long, _
            ByRef sToolTipText As String, _
            ByVal bAutosize As Boolean, _
            ByVal iItemData As Long, _
            ByVal bEnabled As Boolean, _
            ByVal bVisible As Boolean, _
            ByRef vButtonBefore As Variant) _
                As cButton
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Add a button to the collection.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        
        Dim lsText As String
        Dim liIndex As Long
        Dim bSucceed As Boolean
        
        If LenB(sKey) Then
            If pButtons_FindKey(sKey) <> NegOneL Then gErr vbccKeyAlreadyExists, cButtons
        End If
        
        If Not IsMissing(vButtonBefore) Then
            liIndex = pButtons_GetIndex(vButtonBefore)
            If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cButtons
        Else
            liIndex = miButtonCount
        End If
        
        If LenB(sText) Then
            lsText = sText & vbNullChar
        Else
            lsText = vbNullString
        End If
        
        With mtButton
            
            If Not moTrackMenu Is Nothing Then
                iStyle = tbarButtonDropDown
            End If
            
            .fsStyle = (iStyle And &HFF&) Or _
                       (-bAutosize * BTNS_AUTOSIZE)

            
            .fsState = (-bEnabled * TBSTATE_ENABLED) Or _
                       ((bVisible + OneL) * TBSTATE_HIDDEN) 'Or _
                       (-CBool(CBool(GetWindowLong(mhWnd, GWL_STYLE) And CCS_VERT) And moRebar Is Nothing) * TBSTATE_WRAP)
            
            .idCommand = NextItemIdShort()
            
            If iStyle <> tbarButtonSeparator Then
                .iString = StrPtr(lsText)
                .iBitmap = iIconIndex And &HFFFF&
            Else
                .iBitmap = iIconIndex 'this is actually the width of the separator (0 for default)
                .iString = ZeroL
            End If
            
            .dwData = iItemData
            
        End With
        
        If miButtonCount = liIndex Then
            bSucceed = SendMessage(mhWnd, TB_ADDBUTTONSW, OneL, VarPtr(mtButton))
        Else
            bSucceed = SendMessage(mhWnd, TB_INSERTBUTTONW, liIndex, VarPtr(mtButton))
        End If
        
        Debug.Assert bSucceed
        
        If bSucceed Then
            If miButtonCount >= miButtonUbound Then
                Dim liNewUbound As Long
                liNewUbound = RoundToInterval(miButtonCount)
                ReDim Preserve mtButtons(0 To liNewUbound)
                miButtonUbound = liNewUbound
            End If
            
            If liIndex < miButtonCount Then
                With mtButtons(miButtonCount)
                    .sKey = vbNullString
                    .sToolTip = vbNullString
                End With
                CopyMemory mtButtons(liIndex + OneL).iId, mtButtons(liIndex).iId, (miButtonCount - liIndex) * LenB(mtButtons(0))
                ZeroMemory mtButtons(liIndex).iId, LenB(mtButtons(0))
            End If
            
            With mtButtons(liIndex)
                .sKey = StrConv(sKey & vbNullChar, vbFromUnicode)
                .sToolTip = StrConv(sToolTipText & vbNullChar, vbFromUnicode)
                .iId = mtButton.idCommand
                Set fButtons_Add = New cButton
                fButtons_Add.fInit Me, .iId, liIndex
            End With
            
            miButtonCount = miButtonCount + OneL
            
            pResize
            
            If Not moRebar Is Nothing Then moRebar.fToolbar_Resize Me, miRebarId
            
        End If
    End If
End Function

Friend Sub fButtons_Remove(ByRef vButton As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove a button from the collection.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        Dim liIndex As Long
        liIndex = pButtons_GetIndex(vButton)
        If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, cButtons
        If SendMessage(mhWnd, TB_DELETEBUTTON, liIndex, ZeroL) Then
            pDeleteButton liIndex
            If Not moRebar Is Nothing Then moRebar.fToolbar_Resize Me, miRebarId
        End If
    End If
End Sub

Friend Property Get fButtons_Exists(ByRef vButton As Variant) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the given button exists in the collection.
'---------------------------------------------------------------------------------------
    fButtons_Exists = (pButtons_GetIndex(vButton) <> NegOneL)
End Property

Friend Sub fButtons_Clear()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove all buttons from the collection.
'---------------------------------------------------------------------------------------
    miButtonCount = ZeroL
    If mhWnd Then
        Dim i As Long
        For i = SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) To ZeroL Step NegOneL
            SendMessage mhWnd, TB_DELETEBUTTON, i, ZeroL
        Next
    End If
    pResize
    If Not moRebar Is Nothing Then moRebar.fToolbar_Resize Me, miRebarId
End Sub

Friend Property Get fButtons_Count() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the count of the buttons in the collection.
'---------------------------------------------------------------------------------------
    fButtons_Count = miButtonCount
    Debug.Assert SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) = fButtons_Count
End Property

Friend Property Get fButtons_Item(ByRef vButton As Variant) As cButton
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the specified cButton object.
'---------------------------------------------------------------------------------------
    
    Set fButtons_Item = pButton(pButtons_GetIndex(vButton))
    If fButtons_Item Is Nothing Then gErr vbccKeyOrIndexNotFound, cButtons
    
End Property


Private Function pButtons_GetIndex(ByRef vButton As Variant) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the index of a button given its key, index or object.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    If VarType(vButton) = vbObject Then
        Dim loButton As cButton
        Set loButton = vButton
        If Not loButton.fIsMine(Me) Then GoTo handler
        pButtons_GetIndex = loButton.Index
    ElseIf VarType(vButton) = vbString Then
        pButtons_GetIndex = pButtons_FindKey(CStr(vButton))
    Else
        pButtons_GetIndex = CLng(vButton) - OneL
        If pButtons_GetIndex < ZeroL Or pButtons_GetIndex >= miButtonCount Then GoTo handler
    End If
    On Error GoTo 0
    Exit Function
handler:
    On Error GoTo 0
    pButtons_GetIndex = NegOneL
End Function

Private Function pButtons_FindKey(ByRef sKey As String) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the index of a button given its key.
'---------------------------------------------------------------------------------------
    Dim ls As String
    Dim liPtr As Long
    If LenB(sKey) Then
        ls = StrConv(sKey & vbNullChar, vbFromUnicode)
        liPtr = StrPtr(ls)
        For pButtons_FindKey = ZeroL To miButtonCount - OneL
            If LenB(mtButtons(pButtons_FindKey).sKey) Then
                If lstrcmp(liPtr, StrPtr(mtButtons(pButtons_FindKey).sKey)) = ZeroL Then Exit Function
            End If
        Next
    End If
    pButtons_FindKey = NegOneL
End Function

Friend Sub fRebar_Attach(ByVal oRebar As ucRebar, ByVal iId As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Attach the toolbar to a rebar.
'---------------------------------------------------------------------------------------
    miRebarId = iId
    Set moRebar = oRebar
    pSetStyle
End Sub

Friend Sub fRebar_Detach()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Detach the toolbar from a rebar.
'---------------------------------------------------------------------------------------
    Set moRebar = Nothing
    miRebarId = ZeroL
    pResize
End Sub

Friend Property Get fRebar_cxIdeal() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the ideal length of the rebar band.
'---------------------------------------------------------------------------------------
    Dim tR As RECT
    Dim i As Long
    
    For i = ZeroL To SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) - OneL
        If Not pIsSep(i) Then
            If SendMessage(mhWnd, TB_GETITEMRECT, i, VarPtr(tR)) Then
                If mbRebarVertical Then
                    fRebar_cxIdeal = fRebar_cxIdeal + tR.bottom - tR.Top
                Else
                    fRebar_cxIdeal = fRebar_cxIdeal + tR.Right - tR.Left
                End If
            End If
        End If
    Next
End Property

Friend Property Get fRebar_cyChild() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the ideal depth of the rebar band.
'---------------------------------------------------------------------------------------
    Dim tR As RECT
    Dim i As Long
    
    For i = ZeroL To SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) - OneL
        If Not pIsSep(i) Then
            If SendMessage(mhWnd, TB_GETITEMRECT, i, VarPtr(tR)) Then
                If mbRebarVertical Then
                    If (tR.Right - tR.Left) > fRebar_cyChild Then fRebar_cyChild = (tR.Right - tR.Left)
                Else
                    If (tR.bottom - tR.Top) > fRebar_cyChild Then fRebar_cyChild = (tR.bottom - tR.Top)
                End If
            End If
        End If
    Next
    
    If (miStyle And TBSTYLE_FLAT) = ZeroL Then fRebar_cyChild = fRebar_cyChild + 3&
    
End Property

Friend Property Let fRebar_Vertical(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the rebar is vertical or horizontal.
'---------------------------------------------------------------------------------------
    mbRebarVertical = bNew
    If Not moTrackMenu Is Nothing Then
        Dim i As Long
        For i = ZeroL To miButtonCount - OneL
            If Not pIsSep(i) Then
                pButton_Info(mtButtons(i).iId, TBIF_STYLE) = (pButton_Info(mtButtons(i).iId, TBIF_STYLE) And Not BTNS_AUTOSIZE) Or (BTNS_AUTOSIZE * (bNew + OneL))
            End If
        Next
    End If
    pSetStyle
End Property


Friend Property Get fButton_Text(ByVal iId As Long, ByRef iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the text of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_Text = pButton_Text(iId)
    End If
End Property
Friend Property Let fButton_Text(ByVal iId As Long, ByRef iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the text of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        pButton_Text(iId) = sNew
    End If
End Property

Friend Property Get fButton_IconIndex(ByVal iId As Long, ByRef iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the iconindex of the button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_IconIndex = pButton_Info(iId, TBIF_IMAGE)
    End If
End Property
Friend Property Let fButton_IconIndex(ByVal iId As Long, ByRef iIndex As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the iconindex of the button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        pButton_Info(iId, TBIF_IMAGE) = iNew
    End If
End Property

Friend Property Get fButton_ToolTipText(ByVal iId As Long, ByRef iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the tooltiptext of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_ToolTipText = mtButtons(iIndex).sToolTip
        fButton_ToolTipText = StrConv(LeftB$(fButton_ToolTipText, LenB(fButton_ToolTipText) - OneL), vbUnicode)
    End If
End Property
Friend Property Let fButton_ToolTipText(ByVal iId As Long, ByRef iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the tooltiptext of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        mtButtons(iIndex).sToolTip = StrConv(sNew & vbNullChar, vbFromUnicode)
    End If
End Property

Friend Property Get fButton_Enabled(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the enabled/disable status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_Enabled = pButton_State(iId, TBSTATE_ENABLED)
    End If
End Property
Friend Property Let fButton_Enabled(ByVal iId As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the enabled/disable status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        pButton_State(iId, TBSTATE_ENABLED) = bNew
    End If
End Property

Friend Property Get fButton_Checked(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the checked status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_Checked = pButton_State(iId, TBSTATE_CHECKED)
    End If
End Property
Friend Property Let fButton_Checked(ByVal iId As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the checked status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        'pButton_State(iId, TBSTATE_CHECKED) = bNew
        SendMessage mhWnd, TB_CHECKBUTTON, iId, -bNew
    End If
End Property

Friend Property Get fButton_Pressed(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the pressed status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_Pressed = pButton_State(iId, TBSTATE_PRESSED)
    End If
End Property
Friend Property Let fButton_Pressed(ByVal iId As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the pressed status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        pButton_State(iId, TBSTATE_PRESSED) = bNew
    End If
End Property

Friend Property Get fButton_Visible(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the visible status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_Visible = Not pButton_State(iId, TBSTATE_HIDDEN)
    End If
End Property
Friend Property Let fButton_Visible(ByVal iId As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the visible status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        pButton_State(iId, TBSTATE_HIDDEN) = Not bNew
    End If
End Property

Friend Property Get fButton_Grayed(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the grayed status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_Grayed = pButton_State(iId, TBSTATE_INDETERMINATE)
    End If
End Property
Friend Property Let fButton_Grayed(ByVal iId As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the grayed status of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        pButton_State(iId, TBSTATE_INDETERMINATE) = bNew
    End If
End Property

Friend Property Get fButton_Key(ByVal iId As Long, ByRef iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the key of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_Key = mtButtons(iIndex).sKey
        fButton_Key = StrConv(LeftB$(fButton_Key, LenB(fButton_Key) - OneL), vbUnicode)
    End If
End Property
Friend Property Let fButton_Key(ByVal iId As Long, ByRef iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the key of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        Dim ls As String
        ls = StrConv(sNew & vbNullChar, vbFromUnicode)
        If LenB(sNew) Then
            If pButtons_FindKey(sNew) Then gErr vbccKeyAlreadyExists, cButton
        End If
        mtButtons(iIndex).sKey = ls
    End If
End Property

Friend Property Get fButton_ItemData(ByVal iId As Long, ByRef iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the itemdata for a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_ItemData = pButton_Info(iId, TBIF_LPARAM)
    End If
End Property
Friend Property Let fButton_ItemData(ByVal iId As Long, ByRef iIndex As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the itemdata of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        pButton_Info(iId, TBIF_LPARAM) = iNew
    End If
End Property

Friend Property Get fButton_Left(ByVal iId As Long, ByRef iIndex As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the left coordinate of the button.
'---------------------------------------------------------------------------------------
    Dim tR As RECT
    If pButton_Verify(iId, iIndex) Then
        If pButton_GetRect(iIndex, tR) Then
            fButton_Left = ScaleX((tR.Left), vbPixels, vbContainerPosition) + Extender.Left
        End If
    End If
End Property

Friend Property Get fButton_Width(ByVal iId As Long, ByRef iIndex As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the width of the button.
'---------------------------------------------------------------------------------------
    Dim tR As RECT
    If pButton_Verify(iId, iIndex) Then
        If pButton_GetRect(iIndex, tR) Then
            fButton_Width = ScaleX((tR.Right - tR.Left), vbPixels, vbContainerSize)
        End If
    End If
End Property

Friend Property Get fButton_Top(ByVal iId As Long, ByRef iIndex As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the top coordinate of the button.
'---------------------------------------------------------------------------------------
    Dim tR As RECT
    If pButton_Verify(iId, iIndex) Then
        If pButton_GetRect(iIndex, tR) Then
            fButton_Top = ScaleY((tR.Top), vbPixels, vbContainerPosition) + Extender.Top
        End If
    End If
End Property

Friend Property Get fButton_Height(ByVal iId As Long, ByRef iIndex As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the height of the button.
'---------------------------------------------------------------------------------------
    Dim tR As RECT
    If pButton_Verify(iId, iIndex) Then
        If pButton_GetRect(iIndex, tR) Then
            fButton_Height = ScaleY((tR.bottom - tR.Top), vbPixels, vbContainerSize)
        End If
    End If
End Property

Friend Property Get fButton_Index(ByVal iId As Long, ByRef iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the index of a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        fButton_Index = pCommandToIndex(iId) + OneL
    End If
End Property

Friend Sub fButton_GetIdealPopup(ByVal iId As Long, ByRef iIndex As Long, ByRef fLeft As Single, ByRef fTop As Single, ByRef fExcludeLeft As Single, ByRef fExcludeTop As Single, ByRef fExcludeWidth As Single, ByRef fExcludeHeight As Single, ByRef bPreserveVertAlign As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the ideal settings for a popup menu at a button.
'---------------------------------------------------------------------------------------
    If pButton_Verify(iId, iIndex) Then
        Dim tR As RECT
        Dim tRToolbar As RECT
        If pButton_GetRect(iIndex, tR) Then
            
            Dim lB As Boolean
            
            If moChevron.fVisible Then
                lB = (iIndex >= (SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) - SendMessage(mhWndChevron, TB_BUTTONCOUNT, ZeroL, ZeroL)))
            End If
            
            If lB Then
                bPreserveVertAlign = False
                GetWindowRect GetParent(mhWndChevron), tRToolbar
                If moRebar Is Nothing Then
                    lB = (GetWindowLong(mhWndChevron, GWL_STYLE) And CCS_VERT)
                Else
                    lB = Not mbRebarVertical
                End If
            Else
                bPreserveVertAlign = True
                GetWindowRect mhWnd, tRToolbar
                If moRebar Is Nothing Then
                    lB = (GetWindowLong(mhWnd, GWL_STYLE) And CCS_VERT)
                Else
                    lB = mbRebarVertical
                End If
            End If
            
            If mbVertical Then bPreserveVertAlign = Not bPreserveVertAlign
            If Not moRebar Is Nothing Then
                If mbRebarVertical Then bPreserveVertAlign = Not bPreserveVertAlign
            End If
            MapWindowPoints mhWnd, UserControl.ContainerHwnd, tR, TwoL
            
            fExcludeLeft = ScaleX(tR.Left, vbPixels, vbContainerSize)
            fExcludeTop = ScaleX(tR.Top, vbPixels, vbContainerSize)
            fExcludeWidth = ScaleX(tR.Right - tR.Left, vbPixels, vbContainerSize)
            fExcludeHeight = ScaleX(tR.bottom - tR.Top, vbPixels, vbContainerSize)
            
            'MapWindowPoints ZeroL, UserControl.ContainerHwnd, tRToolbar, TwoL
            'fExcludeLeft = ScaleX(tRToolbar.Left, vbPixels, vbContainerPosition)
            'fExcludeTop = ScaleY(tRToolbar.Top, vbPixels, vbContainerPosition)
            
            'If Not moRebar Is Nothing Then
                'fExcludeWidth = ScaleX(tRToolbar.Right - tRToolbar.Left, vbPixels, vbContainerPosition)
                'fExcludeHeight = ScaleY(tRToolbar.Bottom - tRToolbar.Top, vbPixels, vbContainerPosition)
            'Else
                'fExcludeWidth = ScaleX(Width, vbTwips, vbContainerSize)
                'fExcludeHeight = ScaleX(Height, vbTwips, vbContainerSize)
            'End If
            
            If lB Then
                fLeft = ScaleX(tR.Right, vbPixels, vbContainerPosition)
                fTop = ScaleX(tR.Top, vbPixels, vbContainerPosition)
                
                If fLeft < fExcludeLeft + fExcludeWidth Then fLeft = fExcludeLeft + fExcludeWidth
                
            Else
                fLeft = ScaleX(tR.Left, vbPixels, vbContainerPosition)
                fTop = ScaleX(tR.bottom, vbPixels, vbContainerPosition)
                
                If fTop < fExcludeTop + fExcludeHeight Then fTop = fExcludeTop + fExcludeHeight
                
            End If
        End If
    End If
End Sub

Private Function pButton_Verify(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Verify that a button still exists in the collection.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If iIndex > NegOneL And iIndex < miButtonCount Then
            pButton_Verify = mtButtons(iIndex).iId = iId
        End If
        
        If Not pButton_Verify Then
            iIndex = SendMessage(mhWnd, TB_COMMANDTOINDEX, iId, ZeroL)
            If iIndex > NegOneL And iIndex < miButtonCount Then pButton_Verify = (iId = mtButtons(iIndex).iId)
        End If
        
        If Not pButton_Verify Then
            For iIndex = ZeroL To miButtonCount - OneL
                If mtButtons(iIndex).iId = iId Then Exit For
            Next
            pButton_Verify = (iIndex < miButtonCount)
            'button doesn't exist in the toolbar, but it does in the data structure!
            Debug.Assert Not pButton_Verify
        End If
        
    End If
    If Not pButton_Verify Then gErr vbccItemDetached, cButton
    Debug.Assert SendMessage(mhWnd, TB_COMMANDTOINDEX, iId, ZeroL) = iIndex
End Function

Private Function pButton_GetRect(ByVal iIndex As Long, ByRef tR As RECT) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the bounding rectangle of a button.
'---------------------------------------------------------------------------------------
    Dim lhWnd As Long
    Dim liOffset As Long
    If mhWnd Then
        If moChevron.fVisible Then
            liOffset = SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) - SendMessage(mhWndChevron, TB_BUTTONCOUNT, ZeroL, ZeroL)
            If iIndex >= liOffset Then
                lhWnd = mhWndChevron
                iIndex = iIndex - liOffset
            Else
                lhWnd = mhWnd
            End If
        Else
            lhWnd = mhWnd
        End If
        
        If lhWnd Then
            pButton_GetRect = SendMessage(lhWnd, TB_GETITEMRECT, iIndex, VarPtr(tR))
            If lhWnd = mhWndChevron Then MapWindowPoints lhWnd, mhWnd, tR, 2&
        End If
    End If
End Function

Private Property Get pButton_Text(ByVal iId As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the text of a button.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtButtonInfo
            .cchText = LenB(msTextBuffer)
            .pszText = StrPtr(msTextBuffer)
            .dwMask = TBIF_TEXT
            If SendMessage(mhWnd, TB_GETBUTTONINFO, iId, VarPtr(mtButtonInfo)) > NegOneL Then lstrToStringA .pszText, pButton_Text
        End With
    End If
End Property
Private Property Let pButton_Text(ByVal iId As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the text of a button.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtButtonInfo
            .cchText = LenB(msTextBuffer)
            .pszText = StrPtr(msTextBuffer)
            .dwMask = TBIF_TEXT
            lstrFromStringA .pszText, .cchText, sNew
            SendMessage mhWnd, TB_SETBUTTONINFO, iId, VarPtr(mtButtonInfo)
            If mhWndChevron Then
                If SendMessage(mhWndChevron, TB_COMMANDTOINDEX, iId, ZeroL) > NegOneL Then
                    SendMessage mhWndChevron, TB_SETBUTTONINFO, iId, VarPtr(mtButtonInfo)
                End If
            End If
        End With
    End If
End Property

Private Property Get pButton_Info(ByVal iId As Long, ByVal iMask As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value from the TBBUTTONINFO structure.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtButtonInfo
            .dwMask = iMask
            If SendMessage(mhWnd, TB_GETBUTTONINFO, iId, VarPtr(mtButtonInfo)) > NegOneL Then
                If iMask = TBIF_IMAGE Then
                    pButton_Info = .iImage
                ElseIf iMask = TBIF_LPARAM Then
                    pButton_Info = .lParam
                ElseIf iMask = TBIF_SIZE Then
                    pButton_Info = .cx
                ElseIf iMask = TBIF_STATE Then
                    pButton_Info = .fsState
                ElseIf iMask = TBIF_STYLE Then
                    pButton_Info = .fsStyle
                End If
            End If
        End With
    End If
End Property
Private Property Let pButton_Info(ByVal iId As Long, ByVal iMask As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a value in the TBBUTTONINFO structure.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        With mtButtonInfo
            .dwMask = iMask
            If iMask = TBIF_IMAGE Then
                .iImage = iNew
            ElseIf iMask = TBIF_LPARAM Then
                .lParam = iNew
            ElseIf iMask = TBIF_SIZE Then
                .cx = iNew
            ElseIf iMask = TBIF_STATE Then
                .fsState = iNew
            ElseIf iMask = TBIF_STYLE Then
                .fsStyle = iNew
            End If
            SendMessage mhWnd, TB_SETBUTTONINFO, iId, VarPtr(mtButtonInfo)
            If CBool(mhWndChevron) And iMask = TBIF_STATE Then
                If SendMessage(mhWndChevron, TB_COMMANDTOINDEX, iId, ZeroL) > NegOneL Then
                    SendMessage mhWndChevron, TB_PRESSBUTTON, iId, CBool(iNew And TBSTATE_PRESSED)
                    'sendmessage mhWndChevron, TB_SETBUTTONINFO, iId, VarPtr(mtButtonInfo)
                End If
            End If
        End With
    End If
End Property

Private Property Get pButton_State(ByVal iId As Long, ByVal iState As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether a button has the specified state.
'---------------------------------------------------------------------------------------
    pButton_State = CBool(pButton_Info(iId, TBIF_STATE) And iState)
End Property
Private Property Let pButton_State(ByVal iId As Long, ByVal iState As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether a button has the specified state.
'---------------------------------------------------------------------------------------
    If bNew _
        Then pButton_Info(iId, TBIF_STATE) = (pButton_Info(iId, TBIF_STATE) Or iState) _
        Else pButton_Info(iId, TBIF_STATE) = (pButton_Info(iId, TBIF_STATE) And Not iState)
End Property


Friend Sub fMenu_Track(ByVal iState As eToolMenuTrackState, ByVal iIndex As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Track a button as a top-level menu.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        moChevron.fSuspend (iState = tbarTrackDropped)
        If iState = tbarTrackButtons Then
            SendMessage mhWnd, TB_SETHOTITEM, iIndex, ZeroL
            If mhWndChevron = ZeroL And Not moRebar Is Nothing And iIndex >= pFirstClippedIndex() Then
                miIgnoreExitTrack = miIgnoreExitTrack + OneL
                PostMessage UserControl.hWnd, UM_SHOWCHEVRON, iState, iIndex
            Else
                If mhWndChevron Then
                    iIndex = (iIndex - SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) + SendMessage(mhWndChevron, TB_BUTTONCOUNT, ZeroL, ZeroL))
                    If iIndex < ZeroL Then iIndex = NegOneL
                    SendMessage mhWndChevron, TB_SETHOTITEM, iIndex, ZeroL
                End If
            End If
        ElseIf iState = tbarTrackDropped Then
            SendMessage mhWnd, TB_SETHOTITEM, NegOneL, ZeroL
            If mhWndChevron = ZeroL Then
                If Not (moRebar Is Nothing Or mbTrackFromDropDown) Then
                    If iIndex >= pFirstClippedIndex() Then
                        miIgnoreExitTrack = miIgnoreExitTrack + OneL
                        PostMessage UserControl.hWnd, UM_SHOWCHEVRON, iState, iIndex
                        iIndex = NegOneL
                    End If
                End If
                If mbTrackFromDropDown Then mbTrackFromDropDown = False
            Else
                SendMessage mhWndChevron, TB_SETHOTITEM, NegOneL, ZeroL
            End If
            If iIndex > NegOneL And iIndex < miButtonCount Then
                Dim lbOldPressed As Boolean
                lbOldPressed = pButton_State(mtButtons(iIndex).iId, TBSTATE_PRESSED)
                If Not lbOldPressed Then pButton_State(mtButtons(iIndex).iId, TBSTATE_PRESSED) = True
                Set moDroppedMenuButton = pButton(iIndex)
                UpdateWindow mhWnd
                RaiseEvent ButtonDropDown(moDroppedMenuButton)
                Set moDroppedMenuButton = Nothing
                If Not lbOldPressed Then pButton_State(mtButtons(iIndex).iId, TBSTATE_PRESSED) = False
            End If
        Else
            SendMessage mhWnd, TB_SETHOTITEM, NegOneL, ZeroL
            If mhWndChevron Then SendMessage mhWndChevron, TB_SETHOTITEM, NegOneL, ZeroL
        End If
    End If
End Sub

Friend Property Get fMenu_hWndChevron() As Long
    fMenu_hWndChevron = moChevron.fhWnd
End Property

Friend Sub fMenu_ExitTrack()
    If miIgnoreExitTrack = ZeroL Then RaiseEvent ExitMenuTrack
End Sub

Friend Sub fMenu_SetTimer()
    SetTimer UserControl.hWnd, ZeroL, 25&, ZeroL
End Sub

Friend Sub fMenu_KillTimer()
    KillTimer UserControl.hWnd, ZeroL
End Sub

Friend Sub fMenu_CancelDropDown()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Cancel the drop down menu to track another menu.
'---------------------------------------------------------------------------------------
    Dim lhWnd As Long
    Dim lhWndNext As Long
    
    If mhWnd Then SendMessage mhWnd, TB_SETHOTITEM, NegOneL, ZeroL
    
    lhWnd = UserControl.ContainerHwnd
    
    Do
        lhWndNext = lhWnd
        lhWndNext = GetParent(lhWndNext)
    Loop While CBool(lhWndNext)
    
    If lhWnd Then SendMessage lhWnd, WM_CANCELMODE, ZeroL, ZeroL
    
End Sub

Friend Property Get fMenu_NextVisible(ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the next visible item for arrow key navigation.
'---------------------------------------------------------------------------------------
    Dim liState As Long
    
    If mhWnd Then
        For fMenu_NextVisible = iIndex + OneL To miButtonCount - OneL
            liState = pButton_Info(mtButtons(fMenu_NextVisible).iId, TBIF_STATE)
            If (liState And TBSTATE_HIDDEN) = ZeroL And (liState And TBSTATE_ENABLED) = TBSTATE_ENABLED Then Exit For
        Next
        If fMenu_NextVisible >= miButtonCount Then
            For fMenu_NextVisible = ZeroL To iIndex - OneL
                If pButton_State(mtButtons(fMenu_NextVisible).iId, TBSTATE_HIDDEN) = False Then Exit For
            Next
        End If
    End If
End Property

Friend Property Get fMenu_PrevVisible(ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the previous visible item for arrow key navigation.
'---------------------------------------------------------------------------------------
    Dim liState As Long
    If mhWnd Then
        For fMenu_PrevVisible = iIndex - OneL To ZeroL Step NegOneL
            liState = pButton_Info(mtButtons(fMenu_PrevVisible).iId, TBIF_STATE)
            If (liState And TBSTATE_HIDDEN) = ZeroL And (liState And TBSTATE_ENABLED) = TBSTATE_ENABLED Then Exit For
        Next
        If fMenu_PrevVisible < ZeroL Then
            For fMenu_PrevVisible = miButtonCount - OneL To iIndex + OneL Step NegOneL
                If pButton_State(mtButtons(fMenu_PrevVisible).iId, TBSTATE_HIDDEN) = False Then Exit For
            Next
        End If
    End If
End Property

Friend Property Get fMenu_ChevronHitTest(ByRef tP As POINT) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the given point is over the toolbar's
'             chevron if it is on a rebar.
'---------------------------------------------------------------------------------------
    If Not moRebar Is Nothing Then
        fMenu_ChevronHitTest = moRebar.fToolbar_ChevronHitTest(Me, miRebarId, tP)
        
        If fMenu_ChevronHitTest And Not moChevron.fVisible Then
            fMenu_CancelDropDown
            miIgnoreExitTrack = miIgnoreExitTrack + OneL
            PostMessage UserControl.hWnd, UM_SHOWCHEVRON, moTrackMenu.fTrackState, moTrackMenu.fTrackIndex
            moTrackMenu.fTrack tbarTrackNone, NegOneL
        End If
    End If
End Property

Friend Property Get fMenu_HitTest(ByRef tP As POINT) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the item at the given coordinates.
'---------------------------------------------------------------------------------------
    Dim ltP As POINT
    LSet ltP = tP
    
    fMenu_HitTest = NegOneL
    If mhWnd Then
        If WindowFromPoint(ltP.x, ltP.y) = mhWnd Then
            ScreenToClient mhWnd, ltP
            fMenu_HitTest = SendMessage(mhWnd, TB_HITTEST, ZeroL, VarPtr(ltP))
            If Not (fMenu_HitTest < miButtonCount) Then fMenu_HitTest = NegOneL
        ElseIf moChevron.fVisible Then
            If WindowFromPoint(ltP.x, ltP.y) = mhWndChevron Then
                ScreenToClient mhWndChevron, ltP
                fMenu_HitTest = SendMessage(mhWndChevron, TB_HITTEST, ZeroL, VarPtr(ltP))
                If fMenu_HitTest >= ZeroL Then fMenu_HitTest = fMenu_HitTest + (SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) - SendMessage(mhWndChevron, TB_BUTTONCOUNT, ZeroL, ZeroL))
            End If
        End If
    End If
    If fMenu_HitTest >= miButtonCount Or fMenu_HitTest < NegOneL Then fMenu_HitTest = NegOneL
    If fMenu_HitTest <> NegOneL Then
        If Not pButton_State(mtButtons(fMenu_HitTest).iId, TBSTATE_ENABLED) Then fMenu_HitTest = NegOneL
    End If
End Property

Friend Sub fMenu_PostTrack(ByVal iState As eToolMenuTrackState, ByVal iIndex As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Post a message to the queue which will cause is to track a menu.
'---------------------------------------------------------------------------------------
    PostMessage UserControl.hWnd, UM_TRACKMENU, iState, iIndex
End Sub

Friend Property Get fMenu_MapAccelerator(ByVal iChar As Integer, ByVal iStart As Long, ByRef bDuplicate As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Map character input to buttons.
'---------------------------------------------------------------------------------------
    fMenu_MapAccelerator = pMapAccelerator(iChar, iStart, bDuplicate)
End Property

Friend Property Let fMenu_IgnoreMenuKeyPress(ByVal bNew As Boolean)
    mbIgnoreMenuKey = bNew
End Property

Friend Sub fTrackKey_KeyUp(ByVal bNonTrackedKeyPressed As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Alt or F10 has been released without any other keys pressed while
'             it was held down.  Enter menu tracking mode.
'---------------------------------------------------------------------------------------
    If bNonTrackedKeyPressed = False And (Not (moTrackMenu Is Nothing Or mhWnd = ZeroL) And miButtonCount > ZeroL) Then
        Debug.Assert moTrackMenu.fTrackState = tbarTrackNone
        If moTrackMenu.fTrackState = tbarTrackNone Then
            Dim liIndex As Long
            liIndex = SendMessage(mhWnd, TB_GETHOTITEM, ZeroL, ZeroL)
            If liIndex < ZeroL Then liIndex = ZeroL
            fMenu_PostTrack tbarTrackButtons, liIndex
        End If
    End If
End Sub

Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the font used by the toolbar.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property

Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the font used by the toolbar.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
        Else Set moFont = oNew
    pSetFont
    pPropChanged PROP_Font
End Property

Public Property Get ImageList() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the imagelist used by the toolbar.
'---------------------------------------------------------------------------------------
    Set ImageList = moImageList
End Property
Public Property Set ImageList(ByVal oNew As cImageList)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the imagelist used by the toolbar.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Set moImageList = Nothing
    Set moImageListEvent = Nothing
    Set moImageList = oNew
    Set moImageListEvent = oNew
    On Error Resume Next
    miImageSource = tbarDefImageCustom
    pSetImageList mhWnd
End Property

Public Property Get Buttons() As cButtons
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a collection of buttons.
'---------------------------------------------------------------------------------------
    Set Buttons = New cButtons
    Buttons.fInit Me
End Property

Public Property Get List() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the buttons use a list style, showing the
'             text to the right of the icon instead of below it.
'---------------------------------------------------------------------------------------
    List = miStyle And TBSTYLE_LIST
End Property
Public Property Let List(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the buttons use a list style, showing the
'             text to the right of the icon instead of below it.
'---------------------------------------------------------------------------------------
    If mhWnd Then ShowWindow mhWnd, SW_HIDE
    
    pSetStyle TBSTYLE_LIST * (-bNew), TBSTYLE_LIST * (bNew + 1)
    pPropChanged PROP_Style
    pResize
    
    If mhWnd Then ShowWindow mhWnd, SW_SHOWNORMAL
End Property

Public Property Get Wrapable() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether a toolbar will wrap if there is not enough
'             room for the buttons.
'---------------------------------------------------------------------------------------
    Wrapable = miStyle And TBSTYLE_WRAPABLE
End Property
Public Property Let Wrapable(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether a toolbar will wrap if there is not enough
'             room for the buttons.
'---------------------------------------------------------------------------------------
    If mhWnd Then ShowWindow mhWnd, SW_HIDE
    
    pSetStyle TBSTYLE_WRAPABLE * (-bNew), TBSTYLE_WRAPABLE * (bNew + 1)
    pPropChanged PROP_Style
    
    pResize
    
    If mhWnd Then ShowWindow mhWnd, SW_SHOWNORMAL
    
End Property

Public Property Get Flat() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the buttons are flat instead of 3d.
'---------------------------------------------------------------------------------------
    Flat = miStyle And TBSTYLE_FLAT
End Property
Public Property Let Flat(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a value indicating whether the buttons are flat instead of 3d.
'---------------------------------------------------------------------------------------
    If mhWnd Then ShowWindow mhWnd, SW_HIDE
    
    pSetStyle TBSTYLE_FLAT * (-bNew), TBSTYLE_FLAT * (bNew + 1)
    pPropChanged PROP_Style
    
    If mhWnd Then ShowWindow mhWnd, SW_SHOWNORMAL
End Property

Public Property Get Vertical() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value indicating whether the toolbar is vertical.  This is only
'             meaningful when the alignment is set to vbAlignNone.
'---------------------------------------------------------------------------------------
    Vertical = mbVertical
End Property
Public Property Let Vertical(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a value indicating whether the toolbar is vertical.  This is only
'             meaningful when the alignment is set to vbAlignNone.
'---------------------------------------------------------------------------------------
    If bNew Xor mbVertical Then
        mbVertical = bNew
        pCheckVert
    End If
End Property

Public Sub ContainerKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Trap alt and F10 keypresses to show the menu if in menu mode.
'---------------------------------------------------------------------------------------
    
    If KeyCode = vbKeyMenu Or KeyCode = vbKeyF10 Then
        'If Not mbIgnoreMenuKey Then TrackKey.StartTrack Me, KeyCode
        TrackKey.StartTrack Me, KeyCode
        KeyCode = 0
    Else
        If Shift = vbAltMask Then
            
            Dim liIndex As Long
            Dim liStyle As Long
            Dim liAccel As Long
            
            For liIndex = ZeroL To miButtonCount - OneL
                liAccel = AccelChar(pButton_Text(mtButtons(liIndex).iId))
                If liAccel = KeyCode Then
                    KeyCode = 0
                    If moTrackMenu Is Nothing Then
                        liStyle = pButton_Info(mtButtons(liIndex).iId, TBIF_STYLE)
                        
                        If (liStyle And BTNS_CHECK) = ZeroL Then
                            pButton_State(mtButtons(liIndex).iId, TBSTATE_PRESSED) = True
                            UpdateWindow mhWnd
                            DoEvents
                            Sleep 1
                            If liStyle <> BTNS_WHOLEDROPDOWN Then pButton_State(mtButtons(liIndex).iId, TBSTATE_PRESSED) = False
                        Else
                            If liStyle And BTNS_GROUP _
                                Then SendMessage mhWnd, TB_CHECKBUTTON, mtButtons(liIndex).iId, OneL _
                                Else SendMessage mhWnd, TB_CHECKBUTTON, mtButtons(liIndex).iId, pButton_State(mtButtons(liIndex).iId, TBSTATE_CHECKED) + OneL
                        End If
                        
                        If liStyle = BTNS_WHOLEDROPDOWN Then
                            RaiseEvent ButtonDropDown(pButton(liIndex))
                        Else
                            RaiseEvent ButtonClick(pButton(liIndex))
                        End If
                        
                        If liStyle = BTNS_WHOLEDROPDOWN Then pButton_State(mtButtons(liIndex).iId, TBSTATE_PRESSED) = False
                        
                    Else
                        fMenu_PostTrack tbarTrackDropped, liIndex
                        
                    End If
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Public Property Get MinButtonWidth() As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the minimum button width.
'---------------------------------------------------------------------------------------
    MinButtonWidth = ScaleX(loword(miButtonWidth), vbPixels, vbContainerSize)
End Property
Public Property Let MinButtonWidth(ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the minimum button width.
'---------------------------------------------------------------------------------------
    Dim iNew As Integer
    
    On Error GoTo handler
    iNew = ScaleX(fNew, vbContainerSize, vbPixels)
    If False Then
handler:        iNew = DEF_ButtonWidth And &HFFFF&
    End If
    On Error GoTo 0
    
    miButtonWidth = (miButtonWidth And &HFFFF0000) Or iNew
    If mhWnd Then SendMessage mhWnd, TB_SETBUTTONWIDTH, ZeroL, miButtonWidth
    pPropChanged PROP_ButtonWidth
End Property

Public Property Get MaxButtonWidth() As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the maximum button width.
'---------------------------------------------------------------------------------------
    MaxButtonWidth = ScaleX(hiword(miButtonWidth), vbPixels, vbContainerSize)
End Property
Public Property Let MaxButtonWidth(ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the maximum button width.
'---------------------------------------------------------------------------------------
    Dim iNew As Integer
    
    On Error GoTo handler
    iNew = ScaleX(fNew, vbContainerSize, vbPixels)
    If False Then
handler:        iNew = (DEF_ButtonWidth And &HFFFF0000) \ &H10000
    End If
    On Error GoTo 0
    
    miButtonWidth = (miButtonWidth And &HFFFF&) Or (iNew * &H10000)
    If mhWnd Then SendMessage mhWnd, TB_SETBUTTONWIDTH, ZeroL, miButtonWidth
    pPropChanged PROP_ButtonWidth
End Property

Public Property Get ImageSource() As eToolbarImageSource
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the source of the imagelist in the toolbar.
'---------------------------------------------------------------------------------------
    ImageSource = miImageSource
End Property
Public Property Let ImageSource(ByVal iNew As eToolbarImageSource)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the source of the imagelist in the toolbar.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode Then gErr vbccLetSetNoRunTime, ucToolbar
    miImageSource = iNew
    pSetImageSource
    pSetImageList mhWnd
End Property

Private Sub pSetImageSource()
    If Ambient.UserMode Then
        Select Case miImageSource
        Case tbarDefImageStandardSmall, tbarDefImageViewSmall, tbarDefImageHistorySmall
            Set moImageList = New cImageList
            Set moImageListEvent = moImageList
            moImageList.fCreate 16, 16, SystemColorDepth
            pSetImageList mhWnd
            SendMessage mhWnd, TB_LOADIMAGES, miImageSource, NegOneL
        Case tbarDefImageStandardLarge, tbarDefImageViewLarge, tbarDefImageHistoryLarge
            Set moImageList = New cImageList
            Set moImageListEvent = moImageList
            moImageList.fCreate 24, 24, SystemColorDepth
            pSetImageList mhWnd
            SendMessage mhWnd, TB_LOADIMAGES, miImageSource, NegOneL
        End Select
    End If
End Sub

Public Property Get TextRows() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the number of rows of text that can be displayed in the toolbar.
'---------------------------------------------------------------------------------------
    TextRows = miTextRows
End Property

Public Property Let TextRows(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the number of rows of text that can be displayed in the toolbar.
'---------------------------------------------------------------------------------------
    miTextRows = iNew
    pPropChanged PROP_TextRows
    If mhWnd Then SendMessage mhWnd, TB_SETMAXTEXTROWS, miTextRows, ZeroL
End Property

Public Property Get MenuStyle() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether the toolbar is in menu tracking mode.
'---------------------------------------------------------------------------------------
    MenuStyle = Not moTrackMenu Is Nothing
End Property

Public Property Let MenuStyle(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set a value indicating whether the toolbar is in menu tracking mode.
'---------------------------------------------------------------------------------------
    If Ambient.UserMode Then gErr vbccLetSetNoRunTime, ucToolbar
    Set moTrackMenu = IIf(bNew, New pcTrackToolMenu, Nothing)
    If Not moTrackMenu Is Nothing Then Set moTrackMenu.fOwner = Me
    pPropChanged PROP_MenuStyle
End Property

Public Property Get hWnd() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the hwnd of the usercontrol.
'---------------------------------------------------------------------------------------
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndToolbar() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the hwnd of the toolbar.
'---------------------------------------------------------------------------------------
    If mhWnd Then hWndToolbar = mhWnd
End Property

Public Sub ShowPopup( _
   Optional ByVal fLeft As Single, _
   Optional ByVal fTop As Single, _
   Optional ByVal fExcludeLeft As Single, _
   Optional ByVal fExcludeTop As Single, _
   Optional ByVal fExcludeWidth As Single, _
   Optional ByVal fExcludeHeight As Single, _
   Optional ByVal iPopupPosition As eToolbarPopupPosition, _
   Optional ByVal vFirstItem As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Show a popup toolbar, starting at the first clipped index if no starting
'             index is provided.
'---------------------------------------------------------------------------------------
    
    If mhWndChevron = ZeroL And mhWnd <> ZeroL Then
        
        Dim liIndex As Long
        
        If IsMissing(vFirstItem) _
            Then liIndex = pFirstClippedIndex() _
            Else liIndex = pButtons_GetIndex(vFirstItem)
        
        Debug.Assert liIndex < SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL)
        
        If liIndex < SendMessage(mhWnd, TB_BUTTONCOUNT, ZeroL, ZeroL) Then
            Dim lhWnd As Long
            Dim liStyle As Long
            Dim lbHorizontal As Boolean
            
            lbHorizontal = CBool(iPopupPosition > tbarPopRightUp)
            
            lhWnd = pCreateToolbar()
            
            If lhWnd Then
            
                liStyle = (GetWindowLong(lhWnd, GWL_STYLE) And Not (CCS_LEFT Or TBSTYLE_WRAPABLE Or CCS_NOPARENTALIGN Or CCS_NORESIZE)) _
                           Or (miStyle And TBSTYLE_FLAT) _
                           Or CCS_TOP Or ((lbHorizontal + OneL) * (CCS_LEFT Or TBSTYLE_WRAPABLE Or CCS_NORESIZE))
                
                SetWindowLong lhWnd, GWL_STYLE, liStyle
                SendMessage lhWnd, TB_SETSTYLE, ZeroL, liStyle
                If moTrackMenu Is Nothing Then
                    pCopyButtons mhWnd, lhWnd, liIndex, , -lbHorizontal * TBSTATE_WRAP
                Else
                    pCopyButtons mhWnd, lhWnd, liIndex, , -lbHorizontal * TBSTATE_WRAP, _
                                -lbHorizontal * BTNS_AUTOSIZE, _
                                (lbHorizontal + OneL) * BTNS_AUTOSIZE
                End If
                
                SetWindowPos lhWnd, ZeroL, ZeroL, ZeroL, 20, 33000, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOREDRAW Or SWP_NOZORDER
    
                SendMessage lhWnd, TB_AUTOSIZE, ZeroL, ZeroL
                SendMessage lhWnd, TB_SETPARENT, UserControl.hWnd, ZeroL
                
                Dim tR As RECT
                Dim i As Long
                Dim liWidth As Long
                Dim liHeight As Long
                
                For i = miButtonCount - OneL To ZeroL Step NegOneL
                    If Not pIsSep(i) Then
                        liIndex = SendMessage(lhWnd, TB_COMMANDTOINDEX, mtButtons(i).iId, ZeroL)
                        If liIndex > NegOneL Then
                            If SendMessage(lhWnd, TB_GETITEMRECT, liIndex, VarPtr(tR)) Then
                                If tR.Right > liWidth Then liWidth = tR.Right
                                If tR.bottom > liHeight Then liHeight = tR.bottom
                            End If
                        End If
                    End If
                Next
                
                liStyle = liStyle Or CCS_NORESIZE Or CCS_NOPARENTALIGN
                SetWindowLong lhWnd, GWL_STYLE, liStyle
                SendMessage lhWnd, TB_SETSTYLE, ZeroL, liStyle
                
                With tR
                    .Left = ScaleX(fExcludeLeft, vbContainerPosition, vbPixels)
                    .Right = ScaleX(fExcludeLeft + fExcludeWidth, vbContainerPosition, vbPixels)
                    .Top = ScaleY(fExcludeTop, vbContainerPosition, vbPixels)
                    .bottom = ScaleY(fExcludeTop + fExcludeHeight, vbContainerPosition, vbPixels)
                End With
                
                MapWindowPoints UserControl.ContainerHwnd, ZeroL, tR, TwoL
                
                Dim tP As POINT
                tP.x = ScaleX(fLeft, vbContainerPosition, vbPixels)
                tP.y = ScaleY(fTop, vbContainerPosition, vbPixels)
                
                MapWindowPoints UserControl.ContainerHwnd, ZeroL, tP, OneL
                
                mhWndChevron = lhWnd
                pSubclassToolbar lhWnd, True
                moChevron.fShow Me, RootParent(ContainerHwnd), lhWnd, tP.x, tP.y, liWidth, liHeight, tR, iPopupPosition, pGetChevronThemeable()
                pSubclassToolbar lhWnd, False
                DestroyWindow lhWnd
                mhWndChevron = ZeroL
            End If
        End If
    End If
    
End Sub

Private Function pGetChevronThemeable() As Boolean
    If moRebar Is Nothing _
        Then pGetChevronThemeable = mbThemeable _
        Else pGetChevronThemeable = moRebar.Themeable
End Function

Public Sub HidePopup()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Hide the popup menu if it is displayed.
'---------------------------------------------------------------------------------------
    moChevron.fHide
End Sub

Public Sub SetAlignment(ByVal iNew As evbComCtlAlignment)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the alignment of the control.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Dim vbe As VBControlExtender
    Set vbe = Extender
    vbe.Visible = False
    vbe.Align = iNew
    vbe.Visible = True
    pResize
    On Error GoTo 0
End Sub

Public Property Get Themeable() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Return a value indicating whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    Themeable = mbThemeable
End Property

Public Property Let Themeable(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/22/05
' Purpose   : Set whether the default window theme is to be used if available.
'---------------------------------------------------------------------------------------
    If bNew Xor mbThemeable Then
        pPropChanged PROP_Themeable
        mbThemeable = bNew
    End If
    
    If mhWnd Then EnableWindowTheme mhWnd, mbThemeable
End Property

Public Property Get DroppedMenuButton() As cButton
    Set DroppedMenuButton = moDroppedMenuButton
End Property

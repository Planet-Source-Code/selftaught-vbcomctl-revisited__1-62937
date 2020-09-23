VERSION 5.00
Begin VB.UserControl ucPopupMenus 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "ucPopupMenus.ctx":0000
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   34
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucPopupMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucPopupMenus.ctl        7/15/05
'
'           PURPOSE:
'               Create and manage Win32 popup menus.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Menus/Popup_Menu_ActiveX_DLL/VB6_PopupMenu_DLL_Full_Source.asp
'               cPopupMenu.cls
'
'==================================================================================================

Option Explicit

'####################################
'##   ENUMS                        ##
'####################################

Public Enum ePopupItemStyle
    mnuChecked = 1
    mnuRadioChecked = 2
    mnuDisabled = 4
    mnuSeparator = 8
    mnuDefault = 16
    mnuInvisible = 32
    mnuInfrequent = 64
    mnuRedisplayOnClick = 128
    mnuNewVerticalLine = 256
End Enum

Public Enum ePopupShowFlag
    mnuPreserveHAlign = &H0&
    mnuCenterAlign = &H4&
    mnuVCenterAlign = &H10&
    mnuRightAlign = &H8&
    mnuRightButton = &H2&
    mnuBottomAlign = &H20&
    mnuPreserveVertAlign = &H40&
    mnuNoAnimation = &H4000&
    mnuAnimateLTR = &H400&
    mnuAnimateRTL = &H800&
    mnuAnimateTTB = &H1000&
    mnuAnimateBTT = &H2000&
End Enum

Public Enum ePopupDrawStyle
    mnuGradientHighlight = 1&
    mnuButtonHighlight = 2&
    mnuOfficeXPStyle = 4&
    mnuTitleSeparators = 8&
    mnuImageProcessBitmap = 16&
    mnuShowInfrequent = 32&
End Enum


'####################################
'##   EVENTS                       ##
'####################################

Public Event Click(ByVal oItem As cPopupMenuItem)
Public Event ItemHighlight(ByVal oItem As cPopupMenuItem)
Public Event InitPopupMenu(ByVal oMenu As cPopupMenu)
Public Event UnInitPopupMenu(ByVal oMenu As cPopupMenu)


'####################################
'##   CONSTANTS                    ##
'####################################

Private Const VK_CONTROL As Long = &H11&
Private Const VK_LeftCurlyBracket As Long = &HDB&

Private Const TIMER_Id As Long = &H13267458
Private Const TIMER_Interval As Long = 1

Private Const PROP_Font = "Font"
Private Const PROP_ActiveBackColor = "ABack"
Private Const PROP_ActiveForeColor = "AFore"
Private Const PROP_InactiveBackColor = "Back"
Private Const PROP_InactiveForeColor = "AFore"
Private Const PROP_Flags = "Flags"
Private Const PROP_InfreqShowDelay = "Infreq"

Private Const DEF_ActiveBackColor As Long = vbHighlight
Private Const DEF_ActiveForeColor As Long = vbHighlightText
Private Const DEF_InactiveBackColor As Long = vbMenuBar
Private Const DEF_InactiveForeColor As Long = vbMenuText
Private Const DEF_Flags As Long = mnuImageProcessBitmap Or mnuShowInfrequent
Private Const DEF_InfreqShowDelay As Long = 1500

'####################################
'##   VARIABLES                    ##
'####################################

Private WithEvents moFont               As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage           As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1

Private WithEvents moImageListEvent     As cImageList
Attribute moImageListEvent.VB_VarHelpID = -1
Private moImageList                     As cImageList

Private mbChevronClicked                As Boolean

Private mhMenuChevronHover              As Long
Private miIndexChevronHover             As Long
Private miTickCountChevronHover         As Long

Private mtMenus                         As tMenus

'####################################
'##   EVENTS                       ##
'####################################

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
    Set moFontPage = New pcSupportFontPropPage
End Sub

Private Sub UserControl_InitProperties()
    Set moFont = Font_CreateDefault(Ambient.Font)
    
    With mtMenus
        .iActiveBackColor = DEF_ActiveBackColor
        .iActiveForeColor = DEF_ActiveForeColor
        .iInActiveBackColor = DEF_InactiveBackColor
        .iInActiveForeColor = DEF_InactiveForeColor
        .iFlags = DEF_Flags
        .iInfreqShowDelay = DEF_InfreqShowDelay
        .hWndOwner = RootParent(ContainerHwnd)
    End With

    pMenus_SetFont
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    
    With mtMenus
        .iActiveBackColor = PropBag.ReadProperty(PROP_ActiveBackColor, DEF_ActiveBackColor)
        .iActiveForeColor = PropBag.ReadProperty(PROP_ActiveForeColor, DEF_ActiveForeColor)
        .iInActiveBackColor = PropBag.ReadProperty(PROP_InactiveBackColor, DEF_InactiveBackColor)
        .iInActiveForeColor = PropBag.ReadProperty(PROP_InactiveForeColor, DEF_InactiveForeColor)
        .iFlags = PropBag.ReadProperty(PROP_Flags, DEF_Flags)
        .iInfreqShowDelay = PropBag.ReadProperty(PROP_InfreqShowDelay, DEF_InfreqShowDelay)
        .hWndOwner = RootParent(ContainerHwnd)
    End With

    pMenus_SetFont
End Sub

Private Sub UserControl_Resize()
    SIZE ScaleX(28, vbPixels, vbTwips), ScaleY(28, vbPixels, vbTwips)
End Sub

Private Sub UserControl_Terminate()
    mtMenus.hWndOwner = ZeroL
    If mtMenus.hFont Then moFont.ReleaseHandle mtMenus.hFont
    If mtMenus.hFontBold Then moFont.ReleaseHandle mtMenus.hFontBold
    Set moFontPage = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Font_Write moFont, PropBag, PROP_Font

    With mtMenus
        PropBag.WriteProperty PROP_ActiveBackColor, .iActiveBackColor, DEF_ActiveBackColor
        PropBag.WriteProperty PROP_ActiveForeColor, .iActiveForeColor, DEF_ActiveForeColor
        PropBag.WriteProperty PROP_InactiveBackColor, .iInActiveBackColor, DEF_InactiveBackColor
        PropBag.WriteProperty PROP_InactiveForeColor, .iInActiveForeColor, DEF_InactiveForeColor
        PropBag.WriteProperty PROP_Flags, .iFlags, DEF_Flags
        PropBag.WriteProperty PROP_InfreqShowDelay, .iInfreqShowDelay, DEF_InfreqShowDelay
    End With

End Sub

Private Sub moFontPage_AddFonts(ByVal o As ppFont)
    o.ShowProps PROP_Font
End Sub

Private Sub moFontPage_GetAmbientFont(o As stdole.StdFont)
    Set o = Ambient.Font
End Sub

Private Sub moFont_Changed()
    moFont.OnAmbientFontChanged Ambient.Font
    PropertyChanged PROP_Font
    pMenus_SetFont
End Sub

Private Sub moImageListEvent_Changed()
    pMenus_SetImagelist
End Sub






'####################################
'##   PRIVATE FUNCTIONS            ##
'####################################

Private Function pItem_AllocString(ByRef sString As String) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Allocate an ANSI string to store with a menu item.
'---------------------------------------------------------------------------------------
    
    If LenB(sString) Then
        Dim lsCopy As String
        lsCopy = StrConv(sString & vbNullChar, vbFromUnicode)
        pItem_AllocString = MemAllocFromString(StrPtr(lsCopy), LenB(lsCopy))
        
    End If
    
End Function

Private Function pItem_AllocLString(ByVal lpString As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Get a VB string from a pointer stored with a menu item.
'---------------------------------------------------------------------------------------
    If lpString Then lstrToStringA lpString, pItem_AllocLString
End Function

Private Sub pItem_ForceRemeasure(ByVal pItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Force a menu item to remeasure itself the next time it is shown.
'---------------------------------------------------------------------------------------
    
    Dim lhMenu As Long:  lhMenu = PopupMenu_hMenu(PopupItem_pMenuParent(pItem))
    Dim liId As Long:    liId = PopupItem_Id(pItem)
    Dim liFlags As Long: liFlags = pItem_Flags(pItem)
    
    'reset the MF_OWNERDRAW flag.
    Debug.Assert lhMenu
    If lhMenu Then
        Dim lR As Long
        lR = ModifyMenu(lhMenu, liId, liFlags And Not MF_OWNERDRAW, liId, ByVal ZeroL)
        Debug.Assert lR
        lR = ModifyMenu(lhMenu, liId, liFlags, liId, ByVal pItem)
        Debug.Assert lR
    End If
End Sub

Private Function pItem_FreeString(ByVal lpString As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Free a string that had been stored with a menu item.
'---------------------------------------------------------------------------------------
    If lpString Then
        
        pItem_FreeString = MemFree(lpString)
        
    End If
End Function

Private Function pItem_GetHierarchy(ByVal pMenuRoot As Long, ByVal pMenu As Long, ByVal pItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Allocate a linked list that identifies the hierarchy of items from a top
'             level menu to a given item or menu level.
'---------------------------------------------------------------------------------------

    Do
        If pItem Then pItem_GetHierarchy = Hierarchy_Initialize(PopupItem_Id(pItem), pItem_GetHierarchy)
        If pMenu = pMenuRoot Then Exit Do
        pItem = PopupMenu_pItemParent(pMenu)
        pMenu = PopupItem_pMenuParent(pItem)
    Loop While pItem
    
End Function

Private Function pItem_GetShortcutDisplay(ByVal iShortcutKey As Long, ByVal iShortcutMask As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Get a string to display on the menu item to identify the shortcut, such as
'             "CTRL+Y".
'---------------------------------------------------------------------------------------
    Dim lsKey As String

    Const sAlt = "Alt"
    Const sControl = "Ctrl"
    Const sShift = "Shift"
    Const sPlus = "+"
    
    If CBool(iShortcutMask And (vbCtrlMask Or vbShiftMask Or vbAltMask)) Then

        Select Case iShortcutKey
        Case vbKeyHome:             lsKey = "Home"
        Case vbKeyEnd:              lsKey = "End"
        Case vbKeyLeft:             lsKey = "Left"
        Case vbKeyRight:            lsKey = "Right"
        Case vbKeyUp:               lsKey = "Up"
        Case vbKeyDown:             lsKey = "Down"
        Case vbKeyClear:            lsKey = "Clear"
        Case vbKeyPageUp:           lsKey = "Pg Up"
        Case vbKeyPageDown:         lsKey = "Pg Dn"
        Case vbKeyDelete:           lsKey = "Del"
        Case vbKeyEscape:           lsKey = "Esc"
        Case vbKeyTab:              lsKey = "Tab"
        Case vbKeyReturn:           lsKey = "Return"
        Case vbKeyAdd:              lsKey = "Plus"
        Case vbKeySubtract:         lsKey = "Minus"
        Case vbKeyBack:             lsKey = "Bkspc"
        Case vbKeyDivide:           lsKey = "Divide"
        Case vbKeyMultiply:         lsKey = "Multiply"
        Case vbKeyInsert:           lsKey = "Ins"
        Case vbKeySpace:            lsKey = "Space"
        Case vbKeyF1 To vbKeyF16:   lsKey = "F" & vbKeyF1 - iShortcutKey + OneL
        Case Else:                  lsKey = UCase$(Chr$(iShortcutKey))
        End Select
        
        If CBool(iShortcutMask And vbCtrlMask) _
            Then pItem_GetShortcutDisplay = pItem_GetShortcutDisplay & sControl & sPlus
        If CBool(iShortcutMask And vbShiftMask) _
            Then pItem_GetShortcutDisplay = pItem_GetShortcutDisplay & sShift & sPlus
        If CBool(iShortcutMask And vbAltMask) _
            Then pItem_GetShortcutDisplay = pItem_GetShortcutDisplay & sAlt & sPlus
        
        pItem_GetShortcutDisplay = pItem_GetShortcutDisplay & lsKey
        
    End If

End Function

Private Function pItem_Index(ByVal pItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the 1-based index of a menu item.
'---------------------------------------------------------------------------------------
    Dim lpItem As Long: lpItem = PopupMenu_pItemFirst(PopupItem_pMenuParent(pItem))
    
    Do While CBool(lpItem)
        pItem_Index = pItem_Index + OneL
        If lpItem = pItem Then Exit Do
        lpItem = PopupItem_pItemNext(lpItem)
    Loop
    
    Debug.Assert lpItem
    
End Function

Private Sub pItem_Insert(ByVal pItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Insert a menu item structure in the linked list and the menu itself.
'---------------------------------------------------------------------------------------
    Dim lpItemNext As Long: lpItemNext = PopupItem_pItemNext(pItem)
    Dim lpItemPrev As Long: lpItemPrev = PopupItem_pItemPrev(pItem)
    Dim lpMenu As Long:     lpMenu = PopupItem_pMenuParent(pItem)
    Dim lhMenu As Long:     lhMenu = PopupMenu_hMenu(lpMenu)
    
    If lpItemNext _
        Then PopupItem_pItemPrev(lpItemNext) = pItem _
        Else PopupMenu_pItemLast(lpMenu) = pItem
    
    If lpItemPrev _
        Then PopupItem_pItemNext(lpItemPrev) = pItem _
        Else PopupMenu_pItemFirst(lpMenu) = pItem
    
    If lhMenu = ZeroL Then
        
        Dim lpItemParent As Long
        lpItemParent = PopupMenu_pItemParent(lpMenu)
        
        If lpItemParent Then
            
            Do
                If lhMenu Then DestroyMenu lhMenu
                lhMenu = CreatePopupMenu()
            Loop Until PopupMenus_AddID(lhMenu)
            
            PopupMenu_hMenu(lpMenu) = lhMenu
            
            Dim liId As Long
            liId = PopupItem_Id(lpItemParent)
            PopupItem_Id(lpItemParent) = lhMenu
            PopupMenus_ReleaseId liId
            pItem_Info(PopupMenu_hMenu(PopupItem_pMenuParent(lpItemParent)), liId, MIIM_SUBMENU) = lhMenu
            pItem_ForceRemeasure lpItemParent
            
        Else
            
            lhMenu = CreatePopupMenu()
            PopupMenu_hMenu(lpMenu) = lhMenu
            
        End If
        
    End If
    
    Debug.Assert lhMenu
    
    If lhMenu Then
        
        Dim lR As Long
        
        If lpItemNext Then
            lR = InsertMenu(lhMenu, PopupItem_Id(lpItemNext), _
                                  pItem_Flags(pItem), _
                                  PopupItem_Id(pItem), ByVal pItem)
        Else
            lR = AppendMenu(lhMenu, pItem_Flags(pItem), _
                                  PopupItem_Id(pItem), ByVal pItem)
        End If
        
        Debug.Assert lR
        
    End If
    
    pMenu_IncControl lpMenu
    
End Sub

Private Sub pItem_Remove(ByVal pItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Remove a menu item from the linked list and the menu itself.
'---------------------------------------------------------------------------------------
    Dim lhMenu As Long
    lhMenu = PopupMenu_hMenu(PopupItem_pMenuParent(pItem))
    
    If lhMenu Then
        Dim lR As Long
        lR = RemoveMenu(lhMenu, PopupItem_Id(pItem), MF_BYCOMMAND)
        Debug.Assert lR
    End If
    
    Dim lpMenuChild As Long
    lpMenuChild = PopupItem_pMenuChild(pItem)
    
    If lpMenuChild Then
        PopupMenu_pItemParent(lpMenuChild) = ZeroL
    
        If pMenu_RefCount(lpMenuChild) = ZeroL _
            Then pMenu_RemoveSub lpMenuChild _
            Else pMenu_Insert lpMenuChild
            
    End If
    
    pMenu_IncControl PopupItem_pMenuParent(pItem)
    
    Dim lpItemNext As Long:     lpItemNext = PopupItem_pItemNext(pItem)
    Dim lpItemPrev As Long:     lpItemPrev = PopupItem_pItemPrev(pItem)
    Dim lpMenuParent As Long:   lpMenuParent = PopupItem_pMenuParent(pItem)
    
    If lpItemNext _
        Then PopupItem_pItemPrev(lpItemNext) = lpItemPrev _
        Else PopupMenu_pItemLast(lpMenuParent) = lpItemPrev
    
    If lpItemPrev _
        Then PopupItem_pItemNext(lpItemPrev) = lpItemNext _
        Else PopupMenu_pItemFirst(lpMenuParent) = lpItemNext
    
    PopupItem_pMenuParent(pItem) = ZeroL
    
    If PopupItem_RefCount(pItem) < OneL Then pItem_RemoveSub pItem
    
End Sub

Private Sub pItem_RemoveSub(ByVal pItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Release resources associated with a menu item.
'---------------------------------------------------------------------------------------
    PopupMenus_ReleaseId PopupItem_Id(pItem)
        
    pItem_FreeString PopupItem_lpCaption(pItem)
    pItem_FreeString PopupItem_lpHelp(pItem)
    pItem_FreeString PopupItem_lpKey(pItem)
    pItem_FreeString PopupItem_lpShortcutDisplay(pItem)
    
    PopupItem_Terminate pItem
End Sub

Private Sub pItem_SetStyle(ByVal pItem As Long, ByVal iStyleOr As ePopupItemStyle, ByVal iStyleAndNot As ePopupItemStyle)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Change the style of an item, forcing it to remeasure if necessary.
'---------------------------------------------------------------------------------------
    Const RemeasureStyles As Long = mnuSeparator Or mnuNewVerticalLine Or mnuInvisible
    
    Dim liStyle As ePopupItemStyle:     liStyle = PopupItem_Style(pItem)
    Dim liStyleNew As ePopupItemStyle:  liStyleNew = (liStyle Or iStyleOr) And Not iStyleAndNot
    
    If (liStyle And RemeasureStyles) Xor (liStyle And RemeasureStyles) Then pItem_ForceRemeasure pItem
    
End Sub

Private Function pItem_Verify(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Verify that an item still exists in the parent menu.
'---------------------------------------------------------------------------------------
    pItem_Verify = CBool(PopupItem_pMenuParent(pItem))
    If Not pItem_Verify Then gErr vbccItemDetached, "ucPopupMenus"
End Function


Private Function pMenu_CharMatch(ByVal liChar1 As Long, ByVal liChar2 As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Case insensitive character match.
'---------------------------------------------------------------------------------------
    Const LowercaseOffset As Long = 32
    If (liChar1 <> liChar2) Then
        Select Case liChar1
        Case vbKeyA To vbKeyZ:         liChar2 = liChar2 - LowercaseOffset
        Case vbKeyA - LowercaseOffset _
          To vbKeyZ - LowercaseOffset: liChar2 = liChar2 + LowercaseOffset
        End Select
    End If
    pMenu_CharMatch = liChar1 = liChar2
End Function

Private Sub pMenu_CheckKey(ByVal pMenu As Long, ByRef sKey As String)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Raise an error if the key already exists at this level of the menu.
'---------------------------------------------------------------------------------------
    If pMenu_FindItemHandleByKey(pMenu, sKey) Then gErr vbccKeyAlreadyExists, "ucPopupMenus"
End Sub

Private Sub pMenu_Click(ByVal x As Long, ByVal y As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Force the mouse to click.  This allows us to click a chevron if the mouse
'             hovers over it.
'---------------------------------------------------------------------------------------

   ' mouse_event ABSOLUTE coords run from 0 to 65535:
   x = (x * 65535# / CDbl(Screen.Width \ Screen.TwipsPerPixelX))
   y = (y * 65535# / CDbl(Screen.Height \ Screen.TwipsPerPixelY))
   
   ' Click the mouse:
   mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_ABSOLUTE, x, y, ZeroL, ZeroL
   mouse_event MOUSEEVENTF_LEFTUP Or MOUSEEVENTF_ABSOLUTE, x, y, ZeroL, ZeroL

   ' Move the mouse:
   mouse_event MOUSEEVENTF_MOVE Or MOUSEEVENTF_ABSOLUTE, x + OneL, y + OneL, ZeroL, ZeroL

End Sub

Private Function pMenu_ContainerKeyDown(ByVal pMenu As Long, ByVal iKey As Long, ByVal iMask As evbComCtlKeyboardState) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Recursively search for a match to a keyboard accelerator.
'---------------------------------------------------------------------------------------
    If pMenu Then
        Dim lpItem As Long: lpItem = PopupMenu_pItemFirst(pMenu)
        
        Dim liMask As Long
        Dim liKey As Long
        
        Do While lpItem
            liMask = PopupItem_ShortcutMask(lpItem)
            liKey = PopupItem_ShortcutKey(lpItem)
            
            If CBool(liMask) And CBool(liKey) Then
                If liMask = iMask Then
                    If pMenu_CharMatch(liKey, iKey) Then Exit Do
                End If
            End If
            
            pMenu_ContainerKeyDown = pMenu_ContainerKeyDown(PopupItem_pMenuChild(lpItem), iKey, iMask)
            
            If pMenu_ContainerKeyDown _
                Then Exit Function _
                Else lpItem = PopupItem_pItemNext(lpItem)
            
        Loop
        
        pMenu_ContainerKeyDown = lpItem
        
    End If
End Function

Private Sub pMenu_DestroySidebar(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Destroy the resources associated with a sidebar picture.
'---------------------------------------------------------------------------------------
    Dim lhDc As Long:       lhDc = PopupMenu_SidebarHdc(pMenu)
    Dim lhBmpOld As Long:   lhBmpOld = PopupMenu_SidebarHbmpOld(pMenu)
    Dim lhBmp As Long:      lhBmp = PopupMenu_SidebarHbmp(pMenu)
    
    If lhDc Then
        
        Debug.Assert lhBmpOld
        
        If lhBmpOld Then SelectObject lhDc, lhBmpOld
        
        DeleteObject lhBmp
        DeleteDC lhDc
        
        PopupMenu_SidebarHdc(pMenu) = ZeroL
        PopupMenu_SidebarHbmpOld(pMenu) = ZeroL
        PopupMenu_SidebarHbmp(pMenu) = ZeroL
        PopupMenu_SidebarWidth(pMenu) = ZeroL
        PopupMenu_SidebarHeight(pMenu) = ZeroL
        
    End If
End Sub

Private Function pMenu_FindItemHandle(ByVal pMenu As Long, ByRef vItem As Variant) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Find a menu item.
'---------------------------------------------------------------------------------------
    On Error GoTo catch
    
    Dim liVarType As VbVarType: liVarType = VarType(vItem)
    
    Debug.Assert liVarType = vbLong Or liVarType = vbObject Or liVarType = vbString
    
    Select Case liVarType
    Case vbLong:    pMenu_FindItemHandle = pMenu_FindItemHandleByIndex(pMenu, CLng(vItem))
    Case vbObject:  pMenu_FindItemHandle = pMenu_FindItemHandleByObject(pMenu, vItem)
    Case vbString:  pMenu_FindItemHandle = pMenu_FindItemHandleByKey(pMenu, CStr(vItem))
    End Select
        
    On Error GoTo 0
    Exit Function
    
catch:
    
    Debug.Assert False
    pMenu_FindItemHandle = ZeroL
        
    On Error GoTo 0

End Function

Private Function pMenu_FindItemHandleById(ByVal pMenu As Long, ByVal iId As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Find a menu item from its id.
'---------------------------------------------------------------------------------------
    
    If pMenu Then
        
        Dim lpItem As Long
        lpItem = PopupMenu_pItemFirst(pMenu)
        
        Do While lpItem
            
            If PopupItem_Id(lpItem) = iId _
                Then pMenu_FindItemHandleById = lpItem _
                Else pMenu_FindItemHandleById = pMenu_FindItemHandleById(PopupItem_pMenuChild(lpItem), iId)
            
            If pMenu_FindItemHandleById Then Exit Do
            
            lpItem = PopupItem_pItemNext(lpItem)
            
        Loop
        
    End If
End Function

Private Function pMenu_FindItemHandleByIndex(ByVal pMenu As Long, ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Find a menu item by its index in the collection.
'---------------------------------------------------------------------------------------
    
    If iIndex > ZeroL Then
        
        iIndex = iIndex - OneL
        
        pMenu_FindItemHandleByIndex = PopupMenu_pItemFirst(pMenu)
        
        Do While CBool(iIndex) And CBool(pMenu_FindItemHandleByIndex)
            pMenu_FindItemHandleByIndex = PopupItem_pItemNext(pMenu_FindItemHandleByIndex)
            iIndex = iIndex - OneL
        Loop
        
        If iIndex Then pMenu_FindItemHandleByIndex = ZeroL
        
    End If
End Function

Private Function pMenu_FindItemHandleByKey(ByVal pMenu As Long, ByRef sKey As String) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Find a menu item by its key.
'---------------------------------------------------------------------------------------
    
    If LenB(sKey) Then
        
        Dim lsKey As String
        Dim lpKey As Long
        Dim lpKeyCmp As Long
        
        lsKey = StrConv(sKey & vbNullChar, vbFromUnicode)
        lpKey = StrPtr(lsKey)
        
        pMenu_FindItemHandleByKey = PopupMenu_pItemFirst(pMenu)
        
        Do While pMenu_FindItemHandleByKey
            
            lpKeyCmp = PopupItem_lpKey(pMenu_FindItemHandleByKey)
            
            If lpKeyCmp Then
                If lstrcmp(lpKeyCmp, lpKey) = ZeroL Then Exit Do
            End If
            
            pMenu_FindItemHandleByKey = PopupItem_pItemNext(pMenu_FindItemHandleByKey)
            
        Loop
        
    End If
    
End Function

Private Function pMenu_FindItemHandleByObject(ByVal pMenu As Long, ByVal oItem As Object) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Find a menu item by its cPopupMenuItem object.
'---------------------------------------------------------------------------------------
    
    If TypeOf oItem Is cPopupMenuItem Then
        
        Dim loItem As cPopupMenuItem:   Set loItem = oItem
        Dim lpItem As Long:             lpItem = loItem.fpItem
        
        If PopupItem_pMenuParent(lpItem) = pMenu _
            Then pMenu_FindItemHandleByObject = lpItem
        
    End If
    
End Function

Private Function pMenu_FindMenuHandle(ByVal pMenu As Long, ByVal hMenu As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Find a menu or submenu by its handle.
'---------------------------------------------------------------------------------------
    
    If pMenu Then
        
        If PopupMenu_hMenu(pMenu) <> hMenu Then
            
            Dim lpItem As Long: lpItem = PopupMenu_pItemFirst(pMenu)
            Do While CBool(lpItem) And Not CBool(pMenu_FindMenuHandle)
                pMenu_FindMenuHandle = pMenu_FindMenuHandle(PopupItem_pMenuChild(lpItem), hMenu)
                lpItem = PopupItem_pItemNext(lpItem)
            Loop
            
        Else
            pMenu_FindMenuHandle = pMenu
            
        End If
    End If
    
End Function

Private Sub pMenu_ForceRemeasure(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Force all items in a menu to remeasure the next time they are shown.
'---------------------------------------------------------------------------------------
    
    If pMenu Then
    
        Dim lpItem As Long
        
        lpItem = PopupMenu_pItemFirst(pMenu)
        
        Do While lpItem
            
            pItem_ForceRemeasure lpItem
            pMenu_ForceRemeasure PopupItem_pMenuChild(lpItem)
            
            lpItem = PopupItem_pItemNext(lpItem)
            
        Loop
        
    End If
    
End Sub

Private Sub pMenu_GetControlCoords(ByVal oControl As Object, ByVal bShowOnRight As Boolean, ByRef iLeft As Long, ByRef iTop As Long, ByRef tRExclude As RECT)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Get the coordinates to display a menu by a control.
'---------------------------------------------------------------------------------------
    On Error GoTo catch
    GetWindowRect oControl.hWnd, tRExclude
    If bShowOnRight Then
        iLeft = tRExclude.Right
        iTop = tRExclude.Top
    Else
        iLeft = tRExclude.Left
        iTop = tRExclude.bottom
    End If

    On Error GoTo 0
    Exit Sub
catch:

    On Error GoTo catch22
    With tRExclude
        .Left = pMenus_UnscaleX(oControl.Left, True)
        .Right = .Left + pMenus_UnscaleX(oControl.Width, False)
        .Top = pMenus_UnscaleY(oControl.Top, True)
        .bottom = .Top + pMenus_UnscaleY(oControl.Height, False)
    End With

    If bShowOnRight Then
        iLeft = tRExclude.Right
        iTop = tRExclude.Top
    Else
        iLeft = tRExclude.Left
        iTop = tRExclude.bottom
    End If
    On Error GoTo 0
    Exit Sub
catch22:

    On Error GoTo 0
    With tRExclude
        .Left = ZeroL
        .Right = ZeroL
        .Top = ZeroL
        .bottom = ZeroL
    End With
    
    Dim ltCursor As POINT
    GetCursorPos ltCursor
    iLeft = ltCursor.x
    iTop = ltCursor.y
End Sub

Private Function pMenu_GetItemHandle(ByVal pMenu As Long, ByRef vItem As Variant) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Find a menu item, raising an error if it is not found.
'---------------------------------------------------------------------------------------
    
    If Not IsMissing(vItem) Then
        pMenu_GetItemHandle = pMenu_FindItemHandle(pMenu, vItem)
        If pMenu_GetItemHandle = ZeroL Then gErr vbccKeyOrIndexNotFound, "ucPopupMenus"
    End If
End Function

Private Sub pMenu_IncControl(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Increment the control number that will tell us if the menu collection changes
'             during enumeration.
'---------------------------------------------------------------------------------------
    
    Debug.Assert pMenu
    If pMenu Then
        Dim liControl As Long
        liControl = PopupMenu_Control(pMenu)
        Incr liControl
        PopupMenu_Control(pMenu) = liControl
    End If
End Sub

Private Function pMenu_InitPopup(ByVal lpMenu As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Add a chevron if necessary, inject keys if an item is being redisplayed,
'             and raise the InitPopupMenu and UnInitPopupMenu events.
'---------------------------------------------------------------------------------------
    
    If Not CBool(lParam And &HFFFF0000) Then
        
        Dim lpMenuPopup As Long
        lpMenuPopup = pMenu_FindMenuHandle(lpMenu, wParam)
        Debug.Assert CBool(lpMenuPopup) And CBool(iMsg = WM_INITMENUPOPUP Or iMsg = WM_UNINITMENUPOPUP)
        
        If lpMenuPopup Then
            
            If iMsg = WM_INITMENUPOPUP Then
                
                RaiseEvent InitPopupMenu(pMenu(lpMenuPopup))
                
                If Not CBool(mtMenus.iFlags And mnuShowInfrequent) Then
                    
                    If lpMenuPopup Then
                        
                        Dim lpItem As Long: lpItem = PopupMenu_pItemFirst(lpMenuPopup)
                        Do While lpItem
                            If PopupItem_Style(lpItem) And mnuInfrequent Then Exit Do
                            lpItem = PopupItem_pItemNext(lpItem)
                        Loop
                        
                        If lpItem Then
                            lpItem = PopupMenu_pItemLast(lpMenuPopup)
                            If Not CBool(PopupItem_Style(lpItem) And mnuChevron) Then
                                pItem_Insert PopupItem_Initialize(mnuChevron, PopupMenus_GetID(), _
                                                ZeroL, ZeroL, ZeroL, ZeroL, NegOneL, _
                                                ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, _
                                                lpItem, lpMenuPopup)
                            End If
                        End If

                    End If
                    
                End If
                
                If PopupMenu_pHierarchy(lpMenu) Then
                    keybd_event VK_CONTROL, 0, 0, 0
                    keybd_event VK_LeftCurlyBracket, 0, 0, 0
                    keybd_event VK_LeftCurlyBracket, 0, KEYEVENTF_KEYUP, 0
                    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
                End If
                
            ElseIf iMsg = WM_UNINITMENUPOPUP Then
                
                RaiseEvent UnInitPopupMenu(pMenu(lpMenuPopup))
            
            End If
        End If
    End If
End Function

Private Sub pMenu_Insert(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Insert a menu handle in the linked list.
'---------------------------------------------------------------------------------------
    
    If mtMenus.pMenu Then
        
        Dim lpMenu As Long
        Dim lpMenuNext As Long
        
        lpMenuNext = mtMenus.pMenu
        
        Do While lpMenuNext
            lpMenu = lpMenuNext
            lpMenuNext = PopupMenu_pMenuNext(lpMenu)
        Loop
        
        PopupMenu_pMenuNext(lpMenu) = pMenu
        PopupMenu_pMenuNext(pMenu) = lpMenuNext
        
    Else
        
        mtMenus.pMenu = pMenu
        
    End If
End Sub

Private Function pMenu_MenuSelect(ByVal pMenu As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Raise item highlight events and monitor for the mouse a chevron being selected.
'---------------------------------------------------------------------------------------
    
    Dim liId As Long:   liId = wParam And &HFFFF&
    Dim lhMenu As Long: lhMenu = lParam
    Dim lpMenu As Long: lpMenu = pMenu_FindMenuHandle(pMenu, lhMenu)
    
    miIndexChevronHover = NegOneL
    
    Dim lpItem As Long
    
    If lpMenu Then
        
        lpItem = pMenu_FindItemHandleById(lpMenu, liId)
        If lpItem Then
            RaiseEvent ItemHighlight(pItem(lpItem))
            
            If CBool(PopupItem_Style(lpItem) And mnuChevron) And CBool(mtMenus.iInfreqShowDelay > ZeroL) Then
                miIndexChevronHover = pItem_Index(lpItem) - OneL
                mhMenuChevronHover = lhMenu
                miTickCountChevronHover = GetTickCount()
                SetTimer mtMenus.hWndOwner, TIMER_Id, TIMER_Interval, ZeroL
            End If
        End If
    End If
    
    Static lpItemLastHighlight As Long
    
    If lpItem = ZeroL Then
        If lpItemLastHighlight Then
            lpItemLastHighlight = ZeroL
            RaiseEvent ItemHighlight(Nothing)
        End If
    Else
        lpItemLastHighlight = lpItem
    End If
End Function

Private Function pMenu_MenuChar(ByVal pMenu As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Process character input.  If a key is pressed, search the menu items for
'             items with the key as an accelerator. If one is found, MNC_EXECUTE it.  If more
'             than one is found, MNC_SELECT the first item after the currently highlighted
'             item.  If zero are found, search the menu items for items with the key
'             as the first letter of the text.  If one is found, MNC_EXECUTE, if more than
'             one MNC_SELECT.
'---------------------------------------------------------------------------------------
    Debug.Assert ((wParam And &HFFFF0000) \ &H10000) = MF_POPUP
    If ((wParam And &HFFFF0000) \ &H10000) = MF_POPUP Then
        
        Dim liChar As Long: liChar = wParam And &HFFFF&
        Dim lhMenu As Long: lhMenu = lParam
        
        pMenu_MenuChar = pMenu_ProcessRedisplay(pMenu, lhMenu, liChar)
        
        If pMenu_MenuChar = ZeroL Then
            
            Dim liAccel() As Long
            Dim liFirstChar() As Long
            Dim liAccelCount As Long
            Dim liFirstCharCount As Long
            Dim liPos As Long
            Dim liHilite As Long
            Dim liStyle As Long
            Dim lpCaption As Long
            Dim lpItem As Long
            Dim tMI As MENUITEMINFO
            
            Dim liCount As Long: liCount = GetMenuItemCount(lhMenu)
            
            If liCount > ZeroL Then
            
                ReDim liAccel(0 To liCount - OneL)
                ReDim liFirstChar(0 To liCount - OneL)
                
                liHilite = NegOneL
                
                tMI.cbSize = LenB(tMI)
                tMI.fMask = MIIM_STATE Or MIIM_DATA
                
                For liPos = ZeroL To liCount - OneL
                    If GetMenuItemInfo(lhMenu, liPos, True, tMI) Then
                        If CBool(tMI.fState And MF_HILITE) Then liHilite = liPos
                        lpItem = tMI.dwItemData
                        liStyle = PopupItem_Style(lpItem)
                        lpCaption = PopupItem_lpCaption(lpItem)
                        
                        If Not CBool(liStyle And mnuDisabled) _
                           And Not CBool(liStyle And mnuInvisible) _
                           And Not CBool(liStyle And mnuSeparator) _
                           And (CBool(mtMenus.iFlags And mnuShowInfrequent) Or Not CBool(liStyle And mnuInfrequent)) Then         'CBool(tMI.fState And MF_ENABLED) And Not CBool(tMI.fState And MF_SEPARATOR) Then
                            
                            If pMenu_CharMatch(PopupItem_Accelerator(lpItem), liChar) Then
                                liAccel(liAccelCount) = liPos
                                liAccelCount = liAccelCount + OneL
                            End If
                            
                            If lpCaption Then
                                If pMenu_CharMatch(CLng(MemOffset16(lpCaption, ZeroL) And &HFF&), liChar) Then
                                    liFirstChar(liFirstCharCount) = liPos
                                    liFirstCharCount = liFirstCharCount + OneL
                                End If
                            End If
                        End If
                        
                    Else
                        Debug.Assert False
                    End If
        
                Next
                
                If liAccelCount > ZeroL Then
                    If liHilite > NegOneL And liAccelCount > OneL Then
                        For liPos = ZeroL To liAccelCount - OneL
                            pMenu_MenuChar = liAccel(liPos)
                            If pMenu_MenuChar > liHilite Then Exit For
                        Next
                        If liPos = liAccelCount Then pMenu_MenuChar = liAccel(0)
                    Else
                        pMenu_MenuChar = liAccel(0)
                    End If
        
                    If liAccelCount = OneL _
                        Then pMenu_MenuChar = pMenu_MenuChar Or (MNC_EXECUTE * &H10000) _
                        Else pMenu_MenuChar = pMenu_MenuChar Or (MNC_SELECT * &H10000)
        
                ElseIf liFirstCharCount > ZeroL Then
                    If liHilite > NegOneL And liFirstCharCount > OneL Then
                        For liPos = ZeroL To liFirstCharCount - OneL
                            pMenu_MenuChar = liFirstChar(liPos)
                            If liFirstChar(liPos) > liHilite Then Exit For
                        Next
                        If liPos = liFirstCharCount Then pMenu_MenuChar = liFirstChar(0)
                    Else
                        pMenu_MenuChar = liFirstChar(0)
                    End If
        
                    If liFirstCharCount = OneL _
                        Then pMenu_MenuChar = pMenu_MenuChar Or (MNC_EXECUTE * &H10000) _
                        Else pMenu_MenuChar = pMenu_MenuChar Or (MNC_SELECT * &H10000)
                    
                End If
            End If
        End If
    End If
End Function

Private Function pMenu_ProcessRedisplay(ByVal pMenu As Long, ByVal hMenu As Long, ByVal iChar As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : If a keystroke was injected to redisplay a specific item, return a value
'             asking the TrackPopupMenuEx function to display it.
'---------------------------------------------------------------------------------------
    Dim lpHierarchy As Long: lpHierarchy = PopupMenu_pHierarchy(pMenu)
    
    'Debug.Assert lpHierarchy = ZeroL Or iChar = VK_LeftCurlyBracket
    'CBool(iChar = VK_LeftCurlyBracket) And
    
    If CBool(lpHierarchy) Then
        
        Dim liId As Long
        liId = Hierarchy_iId(lpHierarchy)
        
        Dim i As Long
        Dim tMI As MENUITEMINFO
        Dim liCount As Long
        
        liCount = GetMenuItemCount(hMenu)
        
        tMI.cbSize = LenB(tMI)
        tMI.fMask = MIIM_ID
        
        For i = ZeroL To liCount - OneL
            If GetMenuItemInfo(hMenu, i, True, tMI) Then
                If tMI.wID = liId Then Exit For
            Else
                Debug.Assert False
            End If
        Next

        Debug.Assert i < liCount
        
        If i < liCount Then
            
            If pItem_Info(hMenu, liId, MIIM_SUBMENU) _
                Then pMenu_ProcessRedisplay = i Or (MNC_EXECUTE * &H10000) _
                Else pMenu_ProcessRedisplay = i Or (MNC_SELECT * &H10000)
            
            Debug.Assert CBool(pItem_Info(hMenu, liId, MIIM_SUBMENU)) Or CBool(Hierarchy_pHierarchyNext(lpHierarchy) = ZeroL)
            
            PopupMenu_pHierarchy(pMenu) = Hierarchy_pHierarchyNext(lpHierarchy)
            Hierarchy_Terminate lpHierarchy
            
        End If
    End If
End Function

Private Function pMenu_RefCount(ByVal pMenu As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the sum of reference counts on the menu structure and all its ancestors
'             and decendents.
'---------------------------------------------------------------------------------------
    Dim lpMenu As Long: lpMenu = pMenu
    Dim lpItemParent As Long
    
    pMenu_RefCount = pMenu_RefCountChildren(pMenu)
    
    Do
        pMenu_RefCount = pMenu_RefCount + PopupMenu_RefCount(lpMenu)
        lpItemParent = PopupMenu_pItemParent(lpMenu)
        If lpItemParent Then lpMenu = PopupItem_pMenuParent(lpItemParent)
    Loop While lpItemParent
    
End Function

Private Function pMenu_RefCountChildren(ByVal pMenu As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the sum of all the reference counts for the decendents of a given menu.
'---------------------------------------------------------------------------------------
    
    Dim lpItem As Long
    Dim lpMenu As Long
    
    lpItem = PopupMenu_pItemFirst(pMenu)
    
    Do While lpItem
        lpMenu = PopupItem_pMenuChild(lpItem)
        
        If lpMenu _
            Then pMenu_RefCountChildren _
                    = pMenu_RefCountChildren + _
                      PopupMenu_RefCount(lpMenu) + _
                      pMenu_RefCountChildren(lpMenu)

        lpItem = PopupItem_pItemNext(lpItem)
    Loop
    
End Function

Private Sub pMenu_Remove(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Remove a menu structure from the linked list.
'---------------------------------------------------------------------------------------
    If pMenu Then
        
        Dim lpMenu As Long
        Dim lpMenuNext As Long
        
        lpMenuNext = mtMenus.pMenu
        
        Do Until lpMenuNext = pMenu Or lpMenuNext = ZeroL
            
            lpMenu = lpMenuNext
            lpMenuNext = PopupMenu_pMenuNext(lpMenu)
            
        Loop
        
        Debug.Assert lpMenuNext
        
        If lpMenuNext Then
            
            If lpMenu _
                Then PopupMenu_pMenuNext(lpMenu) = PopupMenu_pMenuNext(lpMenuNext) _
                Else mtMenus.pMenu = PopupMenu_pMenuNext(lpMenuNext)
            
        End If
        
        pMenu_RemoveSub pMenu

    End If
End Sub

Private Sub pMenu_RemoveChevrons(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Remove chevrons that may have been added to the menu and its decendents.
'---------------------------------------------------------------------------------------
    If pMenu Then
        Dim lpItem As Long
        lpItem = PopupMenu_pItemLast(pMenu)
        
        If lpItem Then
            If PopupItem_Style(lpItem) And mnuChevron Then pItem_Remove lpItem
            
            lpItem = PopupMenu_pItemFirst(pMenu)
            Do While lpItem
                pMenu_RemoveChevrons PopupItem_pMenuChild(lpItem)
                lpItem = PopupItem_pItemNext(lpItem)
            Loop
            
        End If
        
    End If
End Sub

Private Sub pMenu_RemoveHierarchy(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Remove a hierarchy that has been stored for a menu.
'---------------------------------------------------------------------------------------
    Dim lpHierarchy As Long
    Dim lpHierarchyNext As Long
    
    lpHierarchy = PopupMenu_pHierarchy(pMenu)
    
    Debug.Assert lpHierarchy = ZeroL
    
    Do While lpHierarchy
        lpHierarchyNext = Hierarchy_pHierarchyNext(lpHierarchy)
        Hierarchy_Terminate lpHierarchy
        lpHierarchy = lpHierarchyNext
    Loop
        
    PopupMenu_pHierarchy(pMenu) = ZeroL
    
End Sub

Private Sub pMenu_RemoveSub(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Release resources held by the menu structure.
'---------------------------------------------------------------------------------------
    If pMenu Then
        
        Debug.Assert PopupMenu_RefCount(pMenu) = ZeroL
        
        Dim lpItem As Long
        Dim lpItemNext As Long
        
        lpItemNext = PopupMenu_pItemFirst(pMenu)
        
        Do While lpItemNext
            
            lpItem = lpItemNext
            lpItemNext = PopupItem_pItemNext(lpItem)
            
            pItem_Remove lpItem
            
        Loop
        
        Dim lhMenu As Long
        lhMenu = PopupMenu_hMenu(pMenu)
        If lhMenu Then DestroyMenu lhMenu
        
        pMenu_DestroySidebar pMenu
        
        PopupMenu_Terminate pMenu
        
    End If
End Sub

Private Function pMenu_Root(ByVal pMenu As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the topmost menu structure from a given menu.
'---------------------------------------------------------------------------------------
    Dim lpMenu As Long
    Dim lpItemParent As Long
    Dim lpMenuParent As Long
    
    lpMenuParent = pMenu
    
    Do
        
        lpMenu = lpMenuParent
        lpItemParent = PopupMenu_pItemParent(lpMenu)
        
        If lpItemParent Then lpMenuParent = PopupItem_pMenuParent(lpItemParent)
        
    Loop While lpItemParent
    
    pMenu_Root = lpMenu
    
End Function

Private Function pMenu_Timer(ByVal pMenu As Long, ByVal iId As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Check whether the mouse has been over a chevron for more than the alloted time,
'             and click the mouse if it has.
'---------------------------------------------------------------------------------------
    pMenu_Timer = CBool(iId = TIMER_Id)
    If pMenu_Timer Then
        Dim lbContinue As Boolean
        
        If miIndexChevronHover > NegOneL Then
            Dim ltCursor As POINT
            GetCursorPos ltCursor
            
            If MenuItemFromPoint(ZeroL, mhMenuChevronHover, ltCursor.x, ltCursor.y) = miIndexChevronHover Then
                If (GetTickCount() - miTickCountChevronHover) > mtMenus.iInfreqShowDelay _
                    Then pMenu_Click ltCursor.x, ltCursor.y _
                    Else lbContinue = True
            Else
                miTickCountChevronHover = GetTickCount()
                lbContinue = True
            End If
            
        End If
        
        If Not lbContinue Then KillTimer mtMenus.hWndOwner, TIMER_Id
        
    End If
    
End Function


Private Function pMenu_Show( _
            ByVal lpMenu As Long, _
            ByVal ixPixel As Long, _
            ByVal iyPixel As Long, _
            ByVal iFlags As Long, _
            ByRef tRectExclude As RECT) _
                As cPopupMenuItem
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Track a popup menu.
'---------------------------------------------------------------------------------------
    Dim tPM As TPMPARAMS
    Dim lhMenu As Long
    
    lhMenu = PopupMenu_hMenu(lpMenu)
    
    If CBool(IsWindow(mtMenus.hWndOwner)) And CBool(lhMenu) Then
        tPM.cbSize = LenB(tPM)
        
        Dim lpTPM As Long
        If IsRectEmpty(tRectExclude) = ZeroL Then
            LSet tPM.rcExclude = tRectExclude
            lpTPM = VarPtr(tPM)
        End If
        
        If ixPixel < ZeroL Then ixPixel = ZeroL
        If iyPixel < ZeroL Then iyPixel = ZeroL
        
        Dim loSubclass As pcPopupSubclass
        Dim liId As Long
        Dim lpItem As Long
        
        Do
            lpItem = ZeroL
            
            SendMessage mtMenus.hWndOwner, WM_ENTERMENULOOP, OneL, ZeroL
            
            Set loSubclass = New pcPopupSubclass
            loSubclass.fSubclass Me, lpMenu, mtMenus.hWndOwner
            
            liId = TrackPopupMenuEx(lhMenu, iFlags Or TPM_RETURNCMD, ixPixel, iyPixel, mtMenus.hWndOwner, ByVal lpTPM)
            
            loSubclass.fUnSubclass
            Set loSubclass = Nothing
            
            pMenu_RemoveHierarchy lpMenu
            KillTimer mtMenus.hWndOwner, TIMER_Id
            
            If liId Then
                lpItem = pMenu_FindItemHandleById(lpMenu, liId)
                Debug.Assert lpItem
                If lpItem Then
                    If PopupItem_Style(lpItem) And mnuChevron Then
                        mbChevronClicked = True
                        pMenus_DrawStyle(mnuShowInfrequent) = True
                        PopupMenu_pHierarchy(lpMenu) = pItem_GetHierarchy(lpMenu, PopupItem_pMenuParent(lpItem), ZeroL)
                        
                    Else
                        
                        Set pMenu_Show = pItem(lpItem)
                        
                        If CBool(PopupItem_Style(lpItem) And mnuRedisplayOnClick) _
                            Then PopupMenu_pHierarchy(lpMenu) = pItem_GetHierarchy(lpMenu, PopupItem_pMenuParent(lpItem), lpItem) _
                            Else lpItem = ZeroL
                        
                        RaiseEvent Click(pMenu_Show)
                        
                    End If
                End If
            End If
            
            SendMessage mtMenus.hWndOwner, WM_EXITMENULOOP, OneL, ZeroL
            pMenu_RemoveChevrons lpMenu
            
        Loop While lpItem
        
    End If

End Function




Private Sub pMenus_CalcItemHeight()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Calculate the height of a normal item in the given font.
'---------------------------------------------------------------------------------------
    Dim lhDc As Long
    Dim hFontOld As Long
    
    lhDc = PopupMenus_GetDC()
    hFontOld = SelectObject(lhDc, mtMenus.hFont)
    
    If hFontOld Then
        Dim tSize As SIZE
        GetTextExtentPoint32W lhDc, "A", OneL, tSize
        
        With mtMenus
            .iItemHeight = tSize.cy + 4&
            If .iItemHeight < .iIconSize + 6& Then .iItemHeight = .iIconSize + 6&
            'If .iItemHeight < 16& Then .iItemHeight = 16&
        End With
        
        SelectObject lhDc, hFontOld
    End If
    
    PopupMenus_ReleaseDC
    
End Sub

Private Sub pMenus_ForceRemeasure()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Force all menus to remeasure the next time they are shown.
'---------------------------------------------------------------------------------------
    Dim lpMenu As Long
    
    lpMenu = mtMenus.pMenu
    
    Do While lpMenu
        pMenu_ForceRemeasure lpMenu
        lpMenu = PopupMenu_pMenuNext(lpMenu)
    Loop
    
End Sub

Private Sub pMenus_ProcessPicture()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : If image processing is enabled, prepare two progressively lighter versions
'             of the background bitmap.
'---------------------------------------------------------------------------------------
    Dim lhDcSrc As Long
    Dim liSrcWidth As Long
    Dim liSrcHeight As Long

    With mtMenus
        Set .oDcBackgroundLight = Nothing
        Set .oDcBackgroundSuperLight = Nothing
        
        If Not .oDcBackground Is Nothing Then
            
            If CBool(.iFlags And mnuImageProcessBitmap) Then
                With .oDcBackground
                    lhDcSrc = .hDc
                    liSrcWidth = .Width
                    liSrcHeight = .Height
                End With
    
                Dim lhDc As Long
                Dim lhDib As Long
                Dim lhBmpOld As Long
                Dim ltBMP As BITMAPINFO
                Dim lpBits As Long
    
                With ltBMP.bmiHeader
                    .biSize = Len(ltBMP.bmiHeader)
                    .biWidth = liSrcWidth
                    .biHeight = liSrcHeight
                    .biPlanes = 1
                    .biBitCount = 24
                    .biCompression = ZeroL
                    .biSizeImage = ((liSrcWidth * 3 + 3) And &HFFFFFFFC) * liSrcHeight
                End With
    
                lhDc = CreateCompatibleDC(ZeroL)
                If lhDc Then
                    lhDib = CreateDIBSection(lhDc, ltBMP, ZeroL, lpBits, ZeroL, ZeroL)
                    If lhDib Then
                        lhBmpOld = SelectObject(lhDc, lhDib)
                        If lhBmpOld Then
                            BitBlt lhDc, ZeroL, ZeroL, liSrcWidth, liSrcHeight, lhDcSrc, ZeroL, ZeroL, vbSrcCopy
                            
                            pMenus_ProcessPicture_Lighten lpBits, ltBMP
                            
                            Set .oDcBackgroundLight = New pcMemDC
                            With .oDcBackgroundLight
                                .Create liSrcWidth, liSrcHeight
                                BitBlt .hDc, ZeroL, ZeroL, liSrcWidth, liSrcHeight, lhDc, ZeroL, ZeroL, vbSrcCopy
                            End With
                            
                            pMenus_ProcessPicture_Lighten lpBits, ltBMP
                            
                            Set .oDcBackgroundSuperLight = New pcMemDC
                            With .oDcBackgroundSuperLight
                                .Create liSrcWidth, liSrcHeight
                                 BitBlt .hDc, ZeroL, ZeroL, liSrcWidth, liSrcHeight, lhDc, ZeroL, ZeroL, vbSrcCopy
                            End With
                            
                            SelectObject lhBmpOld, lhDc
                            
                        End If
                        DeleteObject lhDib
                    End If
                    DeleteDC lhDc
                End If
            End If
        End If
    End With
End Sub

Private Sub pMenus_ProcessPicture_Lighten(ByVal lpBits As Long, ByRef tBMP As BITMAPINFO)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Lighten the given DIB.
'---------------------------------------------------------------------------------------
    Dim ltSA As SAFEARRAY2D
    Dim lyDib() As Byte

    Dim x As Long, y As Long
    Dim b As Long, G As Long, R As Long
    Dim h As Single, s As Single, L As Single

    With ltSA
        .cbElements = OneL
        .cDims = TwoL
        .Bounds(0).lLbound = ZeroL
        .Bounds(0).cElements = tBMP.bmiHeader.biHeight
        .Bounds(1).lLbound = ZeroL
        .Bounds(1).cElements = (tBMP.bmiHeader.biWidth * 3& + 3&) And &HFFFFFFFC
        .pvData = lpBits
    End With

    SAPtr(lyDib) = VarPtr(ltSA)

    For x = ZeroL To ((tBMP.bmiHeader.biWidth - OneL) * 3&) Step 3&
        For y = ZeroL To tBMP.bmiHeader.biHeight - OneL
            PopupMenus_HLSforRGB lyDib(x + TwoL, y), lyDib(x + OneL, y), lyDib(x, y), h, L, s
            L = L * 6! / 5!
            If (L > 1!) Then L = 1!
            PopupMenus_RGBforHLS h, L, s, R, G, b
            lyDib(x, y) = b
            lyDib(x + OneL, y) = G
            lyDib(x + TwoL, y) = R
        Next
    Next

    SAPtr(lyDib) = ZeroL

End Sub

Private Sub pMenus_SetFont()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Store the new font handles, calculate the new item height and force all
'             items to remeasure.
'---------------------------------------------------------------------------------------
    On Error GoTo handler
    
    Dim loFont As cFont
    
    With mtMenus
        Set loFont = moFont.FontData(fntDataTypeCFont)
        
        If .hFont Then loFont.ReleaseHandle .hFont
        If .hFontBold Then loFont.ReleaseHandle .hFontBold
        
        Dim liOrigWeight As Long
        
        liOrigWeight = loFont.Weight
        loFont.Weight = loFont.Weight + fntWeightLight
        If loFont.Weight > fntWeightHeavy Then loFont.Weight = fntWeightHeavy
        .hFontBold = loFont.GetHandle()
        loFont.Weight = liOrigWeight
        .hFont = loFont.GetHandle()
    End With
    
    PopupMenus_GetDC
    
    pMenus_CalcItemHeight
    pMenus_ForceRemeasure
    
    PopupMenus_ReleaseDC
    
    On Error GoTo 0
    Exit Sub

handler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Sub pMenus_SetDrawStyle(ByVal iStyle As ePopupDrawStyle)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the drawing style of the menus.
'---------------------------------------------------------------------------------------
    Const ForceRemeasureFlags As Long = mnuTitleSeparators Or mnuShowInfrequent Or mnuOfficeXPStyle
    Const ImageProcessFlags As Long = mnuImageProcessBitmap
    
    Dim liOriginalState As Long

    'mask out any invalid flags
    iStyle = iStyle And &H3F&

    liOriginalState = mtMenus.iFlags
    mtMenus.iFlags = iStyle

    If ((liOriginalState And ForceRemeasureFlags) Xor (iStyle And ForceRemeasureFlags)) _
            Then pMenus_ForceRemeasure

    If (liOriginalState And ImageProcessFlags) Xor (iStyle And ImageProcessFlags) _
            Then pMenus_ProcessPicture

End Sub

Private Sub pMenus_SetImagelist()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a new imagelist to be used by the menus.
'---------------------------------------------------------------------------------------
    Dim tR As RECT
    
    With mtMenus
        
        If Not moImageList Is Nothing Then .hIml = moImageList.hIml Else .hIml = ZeroL
        
        If .hIml Then
            ImageList_GetImageRect .hIml, ZeroL, tR
            .iIconSize = tR.bottom - tR.Top
        Else
            .iIconSize = ZeroL
        End If
        
    End With
    
    PopupMenus_GetDC
    
    pMenus_CalcItemHeight
    pMenus_ForceRemeasure
    
    PopupMenus_ReleaseDC
    
End Sub

Private Function pMenus_UnscaleX(ByVal fVal As Single, ByVal bMap As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Scale the given value from container coordinates to pixels, optionally
'             mapping from the container's hwnd to the screen.
'---------------------------------------------------------------------------------------
    Dim tP As POINT
    tP.x = ScaleX(fVal, IIf(bMap, vbContainerPosition, vbContainerSize), vbPixels)
    If bMap Then MapWindowPoints UserControl.ContainerHwnd, ZeroL, tP, OneL
    pMenus_UnscaleX = tP.x
End Function

Private Function pMenus_UnscaleY(ByVal fVal As Single, ByVal bMap As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Scale the given value from container coordinates to pixels, optionally
'             mapping from the container's hwnd to the screen.
'---------------------------------------------------------------------------------------
    Dim tP As POINT
    tP.y = ScaleX(fVal, IIf(bMap, vbContainerPosition, vbContainerSize), vbPixels)
    If bMap Then MapWindowPoints UserControl.ContainerHwnd, ZeroL, tP, OneL
    pMenus_UnscaleY = tP.y
End Function













'####################################
'##   PRIVATE PROPERTIES           ##
'####################################

Private Property Get pItem(ByVal lpItem As Long) As cPopupMenuItem
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a object representing a given item.
'---------------------------------------------------------------------------------------
    Set pItem = New cPopupMenuItem
    pItem.fInit Me, lpItem
End Property

Private Property Get pItem_Flags(ByVal pItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the item flags for a given item.
'---------------------------------------------------------------------------------------
    Dim liStyle As Long
    liStyle = PopupItem_Style(pItem)
    
    Dim lpMenuChild As Long: lpMenuChild = PopupItem_pMenuChild(pItem)
    Dim lhMenuChild As Long: If lpMenuChild Then lhMenuChild = PopupMenu_hMenu(lpMenuChild)
    
    pItem_Flags = (MF_CHECKED * -CBool(liStyle And mnuChecked)) Or _
                  (MF_MENUBARBREAK * -CBool(liStyle And mnuNewVerticalLine)) Or _
                  (MF_SEPARATOR * (-CBool(liStyle And mnuDisabled) Or -CBool(CBool(liStyle And mnuSeparator) Or CBool(liStyle And mnuInvisible) Or (CBool(liStyle And mnuInfrequent) And Not CBool(mtMenus.iFlags And mnuShowInfrequent))))) Or _
                  (MF_POPUP * -CBool(lhMenuChild)) Or _
                   MF_OWNERDRAW
End Property

Private Property Get pItem_Info( _
            ByVal hMenu As Long, _
            ByVal iId As Long, _
            ByVal iMask As Long) _
                As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Query windows for info on a menu item.
'---------------------------------------------------------------------------------------
    Dim tMII As MENUITEMINFO
    tMII.cbSize = Len(tMII)
    tMII.fMask = iMask
    GetMenuItemInfo hMenu, iId, False, tMII
    If iMask = MIIM_STATE Then
        pItem_Info = tMII.fState
    ElseIf iMask = MIIM_ID Then
        pItem_Info = tMII.wID
    ElseIf iMask = MIIM_SUBMENU Then
        pItem_Info = tMII.hSubMenu
    ElseIf iMask = MIIM_DATA Then
        pItem_Info = tMII.dwItemData
    End If
    
End Property

Private Property Let pItem_Info( _
            ByVal hMenu As Long, _
            ByVal iId As Long, _
            ByVal iMask As Long, _
            ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set info on a given menu item.
'---------------------------------------------------------------------------------------
    Dim tMII As MENUITEMINFO
    tMII.cbSize = Len(tMII)
    tMII.fMask = iMask
    If iMask = MIIM_STATE Then
        tMII.fState = iNew
    ElseIf iMask = MIIM_ID Then
        tMII.wID = iNew
    ElseIf iMask = MIIM_SUBMENU Then
        tMII.fMask = tMII.fMask Or MIIM_ID
        tMII.wID = iNew
        tMII.hSubMenu = iNew
    ElseIf iMask = MIIM_DATA Then
        tMII.dwItemData = iNew
    End If
    SetMenuItemInfo hMenu, iId, False, tMII
    
End Property

Private Property Get pItem_Style(ByVal pItem As Long, ByVal iStyle As ePopupItemStyle) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the item has a given style.
'---------------------------------------------------------------------------------------
    pItem_Style = CBool(PopupItem_Style(pItem) And iStyle)
End Property

Private Property Let pItem_Style(ByVal pItem As Long, ByVal iStyle As ePopupItemStyle, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set whether an item uses a given style.
'---------------------------------------------------------------------------------------
    If bNew _
        Then pItem_SetStyle pItem, iStyle, ZeroL _
        Else pItem_SetStyle pItem, ZeroL, iStyle
End Property

Private Property Get pMenu(ByVal lpMenu As Long) As cPopupMenu
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return an object representing a given menu.
'---------------------------------------------------------------------------------------
    Set pMenu = New cPopupMenu
    pMenu.fInit Me, lpMenu
End Property

Private Property Get pMenus_DrawStyle(ByVal iStyle As ePopupDrawStyle) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the given style is applied to menus.
'---------------------------------------------------------------------------------------
    pMenus_DrawStyle = CBool(mtMenus.iFlags And iStyle)
End Property

Private Property Let pMenus_DrawStyle(ByVal iStyle As ePopupDrawStyle, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set whether a given style is applied to menus.
'---------------------------------------------------------------------------------------
    If bNew _
        Then pMenus_SetDrawStyle mtMenus.iFlags Or iStyle _
        Else pMenus_SetDrawStyle mtMenus.iFlags And Not iStyle
End Property






























'####################################
'##   FRIEND FUNCTIONS             ##
'####################################

Friend Sub fItem_AddRef(ByVal pItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Increment the reference count of an item structure so that it cannot be
'             destroyed while there are object references pointing to it.
'---------------------------------------------------------------------------------------
    PopupItem_RefCount(pItem) = PopupItem_RefCount(pItem) + OneL
    
    Dim lpMenuParent As Long
    lpMenuParent = PopupItem_pMenuParent(pItem)
    
    If lpMenuParent Then fMenu_AddRef lpMenuParent
End Sub

Friend Sub fItem_Release(ByVal pItem As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Decrement the reference count of an item.
'---------------------------------------------------------------------------------------
    Dim liRefCount As Long
    liRefCount = PopupItem_RefCount(pItem) - OneL
    PopupItem_RefCount(pItem) = liRefCount
    
    Dim lpMenuParent As Long
    lpMenuParent = PopupItem_pMenuParent(pItem)
    
    If lpMenuParent Then
        fMenu_Release lpMenuParent
    ElseIf liRefCount = ZeroL Then
        pItem_RemoveSub pItem
    End If
    
End Sub

Friend Sub fItem_SetStyle(ByVal pItem As Long, ByVal iStyleOr As ePopupItemStyle, ByVal iStyleAndNot As ePopupItemStyle)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the style of a given item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_SetStyle pItem, iStyleOr, iStyleAndNot
    End If
End Sub



Friend Function fMenu_Add( _
            ByRef pMenu As Long, _
            ByRef sCaption As String, _
            ByRef sHelpString As String, _
            ByRef sKey As String, _
            ByVal iIconIndex As Long, _
            ByVal iStyle As ePopupItemStyle, _
            ByVal iShortcutKey As Integer, _
            ByVal iShortcutMask As evbComCtlKeyboardState, _
            ByVal iItemData As Long, _
            ByRef vItemInsertBefore As Variant) _
                As cPopupMenuItem
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Add a item the the menu collection and return an object that represents
'             the new item.
'---------------------------------------------------------------------------------------

    Const iStyleMask As Long = mnuChecked Or mnuRadioChecked Or mnuDisabled Or mnuSeparator Or _
                         mnuDefault Or mnuInvisible Or mnuInfrequent Or mnuRedisplayOnClick Or mnuNewVerticalLine
    
    pMenu_CheckKey pMenu, sKey
    
    Dim lpItemInsertBefore As Long
    lpItemInsertBefore = pMenu_GetItemHandle(pMenu, vItemInsertBefore)
    
    Dim lpItemInsertAfter As Long
    If lpItemInsertBefore _
        Then lpItemInsertAfter = PopupItem_pItemPrev(lpItemInsertBefore) _
        Else lpItemInsertAfter = PopupMenu_pItemLast(pMenu)
    
    Dim lpItem As Long
    lpItem = PopupItem_Initialize( _
                    iStyle And iStyleMask, PopupMenus_GetID(), _
                    iItemData, iShortcutKey, iShortcutMask, _
                    Asc(LCase$(Chr$(AccelChar(sCaption)))), iIconIndex, _
                    pItem_AllocString(sCaption), pItem_AllocString(sHelpString), _
                    pItem_AllocString(sKey), pItem_AllocString(pItem_GetShortcutDisplay(iShortcutKey, iShortcutMask)), _
                    lpItemInsertBefore, lpItemInsertAfter, pMenu)
                    
    pItem_Insert lpItem
    Set fMenu_Add = pItem(lpItem)
    
End Function

Friend Sub fMenu_AddRef(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Increment the reference count of a menu so that it is not destroyed while
'             there are object references pointing to it.
'---------------------------------------------------------------------------------------
    mPopup.PopupMenu_RefCount(pMenu) = mPopup.PopupMenu_RefCount(pMenu) + OneL
End Sub

Friend Sub fMenu_Clear(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Remove all items from the menu collection.
'---------------------------------------------------------------------------------------
    Dim lpItem As Long
    Dim lpItemNext As Long
    lpItem = PopupMenu_pItemFirst(pMenu)
    
    Do While lpItem
        lpItemNext = PopupItem_pItemNext(lpItem)
        pItem_Remove lpItem
        lpItem = lpItemNext
    Loop
    
End Sub

Friend Sub fMenu_Enum_GetNextItem(ByVal pMenu As Long, tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the next item in an enumeration of the menu collection.
'---------------------------------------------------------------------------------------
    If tEnum.iControl <> PopupMenu_Control(pMenu) Then gErr vbccCollectionChangedDuringEnum, "ucPopupMenus"
    
    If tEnum.iData _
        Then tEnum.iData = PopupItem_pItemNext(tEnum.iData) _
        Else tEnum.iData = PopupMenu_pItemFirst(pMenu)
    
    If tEnum.iData _
        Then Set vNextItem = pItem(tEnum.iData) _
        Else bNoMoreItems = True
    
End Sub

Friend Sub fMenu_Enum_Reset(ByVal pMenu As Long, tEnum As tEnum)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Reset the enumeration data.
'---------------------------------------------------------------------------------------
    tEnum.iControl = PopupMenu_Control(pMenu)
    tEnum.iData = ZeroL
    tEnum.iIndex = ZeroL
End Sub

Friend Sub fMenu_Enum_Skip(ByVal pMenu As Long, tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Skip a given number of items in the enumeration.
'---------------------------------------------------------------------------------------
    If tEnum.iControl <> PopupMenu_Control(pMenu) Then gErr vbccCollectionChangedDuringEnum, "ucPopupMenus"
    
    If iSkipCount > ZeroL Then
        
        Do While iSkipCount
            If tEnum.iData _
                Then tEnum.iData = PopupItem_pItemNext(tEnum.iData) _
                Else tEnum.iData = PopupMenu_pItemFirst(pMenu)
            If tEnum.iData = ZeroL Then Exit Do
            iSkipCount = iSkipCount - OneL
        Loop
        
        bSkippedAll = Not CBool(iSkipCount)
        
    End If
    
End Sub

Friend Sub fMenu_Release(ByVal pMenu As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Decrement the reference count associated with a given menu.
'---------------------------------------------------------------------------------------
    PopupMenu_RefCount(pMenu) = mPopup.PopupMenu_RefCount(pMenu) - OneL
    If pMenu_RefCount(pMenu) = ZeroL Then pMenu_Remove pMenu_Root(pMenu)
    
End Sub

Friend Sub fMenu_Remove(ByVal pMenu As Long, ByRef vItem As Variant)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Remove an item from the menu collection.
'---------------------------------------------------------------------------------------
    pItem_Remove pMenu_GetItemHandle(pMenu, vItem)
End Sub

Friend Sub fMenu_SetSidebar(ByVal pMenu As Long, ByVal oNew As Object)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the sidebar picture displayed by a menu.
'---------------------------------------------------------------------------------------
    pMenu_DestroySidebar pMenu
    
    If Not oNew Is Nothing Then
    
        Dim loMemDC As pcMemDC
        Set loMemDC = New pcMemDC
        loMemDC.CreateFromPicture oNew
        
        Dim liWidth As Long:    liWidth = loMemDC.Width
        Dim liHeight As Long:   liHeight = loMemDC.Height
        Dim lhDc As Long:       lhDc = CreateCompatibleDC(loMemDC.hDc)
        Dim lhBmp As Long:      lhBmp = CreateCompatibleBitmap(loMemDC.hDc, liWidth, liHeight)
        
        PopupMenu_SidebarHdc(pMenu) = lhDc
        PopupMenu_SidebarHbmpOld(pMenu) = SelectObject(lhDc, lhBmp)
        PopupMenu_SidebarHbmp(pMenu) = lhBmp
        PopupMenu_SidebarWidth(pMenu) = liWidth
        PopupMenu_SidebarHeight(pMenu) = liHeight
        
        BitBlt lhDc, ZeroL, ZeroL, liWidth, liHeight, loMemDC.hDc, ZeroL, ZeroL, vbSrcCopy
        
    End If
    
    pMenu_ForceRemeasure pMenu
    
End Sub

Friend Function fMenu_Show( _
            ByVal pMenu As Long, _
            ByVal iFlags As ePopupShowFlag, _
            ByVal fLeft As Single, _
            ByVal fTop As Single, _
   Optional ByVal fExcludeLeft As Single, _
   Optional ByVal fExcludeTop As Single, _
   Optional ByVal fExcludeWidth As Single, _
   Optional ByVal fExcludeHeight As Single) _
                As cPopupMenuItem
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Show the menu at the given coordinates.
'---------------------------------------------------------------------------------------
   
    Dim tRExclude As RECT
    With tRExclude
        .Left = pMenus_UnscaleX(fExcludeLeft, True)
        .Right = .Left + pMenus_UnscaleX(fExcludeWidth, False)
        .Top = pMenus_UnscaleY(fExcludeTop, True)
        .bottom = .Top + pMenus_UnscaleY(fExcludeHeight, False)
    End With
    Set fMenu_Show = pMenu_Show(pMenu, pMenus_UnscaleX(fLeft, True), pMenus_UnscaleY(fTop, True), iFlags, tRExclude)
End Function

Friend Function fMenu_ShowAtControl( _
            ByVal pMenu As Long, _
            ByVal iFlags As ePopupShowFlag, _
            ByVal oControl As Object, _
   Optional ByVal bShowOnRight As Boolean) _
                As cPopupMenuItem
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Show the menu next to a control.
'---------------------------------------------------------------------------------------
    
    Dim tR As RECT
    Dim liLeft As Long, liTop As Long

    pMenu_GetControlCoords oControl, bShowOnRight, liLeft, liTop, tR
    Set fMenu_ShowAtControl = pMenu_Show(pMenu, liLeft, liTop, iFlags, tR)

End Function

Friend Function fMenu_ShowAtCursor( _
            ByVal pMenu As Long, _
            ByVal iFlags As ePopupShowFlag) _
                As cPopupMenuItem
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Show the menu by the cursor.
'---------------------------------------------------------------------------------------
    
    Dim ltCursor As POINT
    GetCursorPos ltCursor
    
    Dim tR As RECT
    Set fMenu_ShowAtCursor = pMenu_Show(pMenu, ltCursor.x, ltCursor.y, iFlags, tR)
End Function




'####################################
'##   FRIEND PROPERTIES            ##
'####################################

Friend Property Get fItem_BreakLine(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the menu item starts on a new vertical line.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_BreakLine = pItem_Style(pItem, mnuNewVerticalLine)
    End If
End Property
Friend Property Let fItem_BreakLine(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the menu item starts on a new vertical line.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuNewVerticalLine) = bNew
    End If
End Property

Friend Property Get fItem_Caption(ByVal pItem As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the caption displayed by the menu item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Caption = pItem_AllocLString(PopupItem_lpCaption(pItem))
    End If
End Property
Friend Property Let fItem_Caption(ByVal pItem As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the caption displayed by the menu item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_FreeString PopupItem_lpCaption(pItem)
        PopupItem_lpCaption(pItem) = pItem_AllocString(sNew)
        pItem_ForceRemeasure pItem
    End If
End Property

Friend Property Get fItem_Checked(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the menu item is checked.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Checked = pItem_Style(pItem, mnuChecked)
    End If
End Property
Friend Property Let fItem_Checked(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the menu item is checked.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuChecked) = bNew
    End If
End Property

Friend Property Get fItem_Default(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the menu item is displayed in bold.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Default = pItem_Style(pItem, mnuDefault)
    End If
End Property
Friend Property Let fItem_Default(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the menu item is displayed in bold.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuDefault) = bNew
    End If
End Property

Friend Property Get fItem_Enabled(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the menu item can be clicked.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Enabled = Not pItem_Style(pItem, mnuDisabled)
    End If
End Property
Friend Property Let fItem_Enabled(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the menu item can be clicked.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuDisabled) = Not bNew
    End If
End Property

Friend Property Get fItem_HelpString(ByVal pItem As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the help string stored by the menu item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_HelpString = pItem_AllocLString(PopupItem_lpHelp(pItem))
    End If
End Property
Friend Property Let fItem_HelpString(ByVal pItem As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the help string stored by the menu item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_FreeString PopupItem_lpHelp(pItem)
        PopupItem_lpHelp(pItem) = pItem_AllocString(sNew)
    End If
End Property

Friend Property Get fItem_IconIndex(ByVal pItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the index of the icon displayed on the item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_IconIndex = PopupItem_IconIndex(pItem)
    End If
End Property
Friend Property Let fItem_IconIndex(ByVal pItem As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the index of the icon displayed on the item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        PopupItem_IconIndex(pItem) = iNew
    End If
End Property

Friend Property Get fItem_Index(ByVal pItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the index of the item in the menu collection.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Index = pItem_Index(pItem)
    End If
End Property

Friend Property Get fItem_Infrequent(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the item is infrequently used.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Infrequent = pItem_Style(pItem, mnuInfrequent)
    End If
End Property
Friend Property Let fItem_Infrequent(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the item is infrequently used.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuInfrequent) = bNew
    End If
End Property

Friend Property Get fItem_ItemData(ByVal pItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value stored by the item for use by the client.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_ItemData = PopupItem_ItemData(pItem)
    End If
End Property
Friend Property Let fItem_ItemData(ByVal pItem As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value stored by the item for use by the client.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        PopupItem_ItemData(pItem) = iNew
    End If
End Property

Friend Property Get fItem_Key(ByVal pItem As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the key stored by the menu item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Key = pItem_AllocLString(PopupItem_lpKey(pItem))
    End If
End Property
Friend Property Let fItem_Key(ByVal pItem As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the key stored by the menu item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pMenu_CheckKey PopupItem_pMenuParent(pItem), sNew
        pItem_FreeString PopupItem_lpKey(pItem)
        PopupItem_lpKey(pItem) = pItem_AllocString(sNew)
    End If
End Property

Friend Property Get fItem_Parent(ByVal pItem As Long) As cPopupMenu
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the menu to which the item belongs.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        Set fItem_Parent = pMenu(PopupItem_pMenuParent(pItem))
    End If
End Property

Friend Property Get fItem_RadioChecked(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the item is checked with a circle instead of a checkmark.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_RadioChecked = pItem_Style(pItem, mnuRadioChecked)
    End If
End Property
Friend Property Let fItem_RadioChecked(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the item is checked with a circle instead of a checkmark.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuRadioChecked) = bNew
    End If
End Property

Friend Property Get fItem_RedisplayOnClick(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the menu redisplays when the item is clicked.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_RedisplayOnClick = pItem_Style(pItem, mnuRedisplayOnClick)
    End If
End Property
Friend Property Let fItem_RedisplayOnClick(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the menu redisplays when the item is clicked.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuRedisplayOnClick) = bNew
    End If
End Property

Friend Property Get fItem_Separator(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the item is a separator.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Separator = pItem_Style(pItem, mnuSeparator)
    End If
End Property
Friend Property Let fItem_Separator(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the item is a separator.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuSeparator) = bNew
    End If
End Property

Friend Property Get fItem_ShortcutDisplay(ByVal pItem As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the string displayed right-aligned on the menu item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_ShortcutDisplay = pItem_AllocLString(PopupItem_lpShortcutDisplay(pItem))
    End If
End Property
Friend Property Let fItem_ShortcutDisplay(ByVal pItem As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the string displayed right-aligned on the menu item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_FreeString PopupItem_lpShortcutDisplay(pItem)
        PopupItem_lpShortcutDisplay(pItem) = pItem_AllocString(sNew)
        pItem_ForceRemeasure pItem
    End If
End Property

Friend Property Get fItem_ShortcutShiftMask(ByVal pItem As Long) As evbComCtlKeyboardState
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the shortcut shift mask for the item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_ShortcutShiftMask = PopupItem_ShortcutMask(pItem)
    End If
End Property
Friend Property Let fItem_ShortcutShiftMask(ByVal pItem As Long, ByVal iNew As evbComCtlKeyboardState)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the shortcut shift mask for the item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        PopupItem_ShortcutMask(pItem) = iNew
        pItem_FreeString PopupItem_lpShortcutDisplay(pItem)
        PopupItem_lpShortcutDisplay(pItem) = pItem_AllocString(pItem_GetShortcutDisplay(PopupItem_ShortcutKey(pItem), iNew))
        pItem_ForceRemeasure pItem
    End If
End Property

Friend Property Get fItem_ShortcutKey(ByVal pItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the shortcut key for the item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_ShortcutKey = PopupItem_ShortcutKey(pItem)
    End If
End Property
Friend Property Let fItem_ShortcutKey(ByVal pItem As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the shortcut key for the item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        PopupItem_ShortcutKey(pItem) = iNew
        pItem_FreeString PopupItem_lpShortcutDisplay(pItem)
        PopupItem_lpShortcutDisplay(pItem) = pItem_AllocString(pItem_GetShortcutDisplay(iNew, PopupItem_ShortcutMask(pItem)))
        pItem_ForceRemeasure pItem
    End If
End Property

Friend Property Get fItem_Style(ByVal pItem As Long) As ePopupItemStyle
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a bitmask identifying the style of the item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Style = PopupItem_Style(pItem)
    End If
End Property
Friend Property Let fItem_Style(ByVal pItem As Long, ByVal iNew As ePopupItemStyle)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a bitmask identifying the style of the item.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_SetStyle pItem, iNew, Not iNew
    End If
End Property

Friend Property Get fItem_SubMenu(ByVal pItem As Long) As cPopupMenu
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a submenu that is the child of the item.
'---------------------------------------------------------------------------------------
    
    Dim lpMenu As Long
    lpMenu = PopupItem_pMenuChild(pItem)
    
    If lpMenu = ZeroL Then
        lpMenu = PopupMenu_Initialize(pItem)
        PopupItem_pMenuChild(pItem) = lpMenu
    End If
    
    Set fItem_SubMenu = New cPopupMenu
    fItem_SubMenu.fInit Me, lpMenu
End Property

Friend Property Get fItem_Visible(ByVal pItem As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the given item appears on the menu.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        fItem_Visible = Not pItem_Style(pItem, mnuInvisible)
    End If
End Property
Friend Property Let fItem_Visible(ByVal pItem As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the given item appears on the menu.
'---------------------------------------------------------------------------------------
    If pItem_Verify(pItem) Then
        pItem_Style(pItem, mnuInvisible) = Not bNew
    End If
End Property




Friend Property Get fMenu_Count(ByVal pMenu As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the number of items in the menu collection.
'---------------------------------------------------------------------------------------
    Dim lpItem As Long
    lpItem = PopupMenu_pItemFirst(pMenu)
    
    Do While lpItem
        fMenu_Count = fMenu_Count + OneL
        lpItem = PopupItem_pItemNext(lpItem)
    Loop
End Property

Friend Property Get fMenu_Control(ByVal pMenu As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the enumeration control number.
'---------------------------------------------------------------------------------------
    fMenu_Control = PopupMenu_Control(pMenu)
End Property

Friend Property Get fMenu_Exists(ByVal pMenu As Long, ByRef vItem As Variant) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the given item exists in the collection.
'---------------------------------------------------------------------------------------
    fMenu_Exists = CBool(pMenu_FindItemHandle(pMenu, vItem))
End Property

Friend Property Get fMenu_Item(ByVal pMenu As Long, ByRef vItem As Variant) As cPopupMenuItem
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return an item from the menu collection.
'---------------------------------------------------------------------------------------
    Set fMenu_Item = pItem(pMenu_GetItemHandle(pMenu, vItem))
End Property

Friend Property Get fMenu_Parent(ByVal lpMenu As Long) As cPopupMenu
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the parent menu of the given menu, if any.
'---------------------------------------------------------------------------------------
    Dim lpItemParent As Long
    
    lpItemParent = PopupMenu_pItemParent(lpMenu)
    If lpItemParent Then
        Set fMenu_Parent = pMenu(PopupItem_pMenuParent(lpMenu))
    End If
End Property

Friend Property Get fMenu_Root(ByVal pMenu As Long) As cPopupMenu
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the root menu.  This can be the same as the given menu.
'---------------------------------------------------------------------------------------
    Set fMenu_Root = New cPopupMenu
    fMenu_Root.fInit Me, pMenu_Root(pMenu)
End Property

Friend Property Get fMenu_SidebarExists(ByVal pMenu As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the menu is drawn with a sidebar.
'---------------------------------------------------------------------------------------
    fMenu_SidebarExists = CBool(PopupMenu_SidebarHdc(pMenu))
End Property

Friend Property Get fMenu_ShowCheckAndIcon(ByVal pMenu As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether space is reserved for a check and an icon.
'---------------------------------------------------------------------------------------
    fMenu_ShowCheckAndIcon = CBool(PopupMenu_Style(pMenu) And mnuShowCheckAndIcon)
End Property

Friend Property Let fMenu_ShowCheckAndIcon(ByVal pMenu As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether space is reserved for a check and an icon.
'---------------------------------------------------------------------------------------
    PopupMenu_Style(pMenu) = mnuShowCheckAndIcon * -bNew
    pMenu_ForceRemeasure pMenu
End Property

Friend Sub fSubclass_Proc(ByVal pMenu As Long, bHandled As Boolean, lReturn As Long, hWnd As Long, iMsg As Long, wParam As Long, lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Handle certain messages sent to the parent window while the menu is active.
'---------------------------------------------------------------------------------------
    Select Case iMsg
    
    Case WM_DRAWITEM:        lReturn = PopupMenus_DrawItem(mtMenus, lParam):    bHandled = CBool(lReturn)
    Case WM_MEASUREITEM:     lReturn = PopupMenus_MeasureItem(mtMenus, lParam): bHandled = CBool(lReturn)
    
    Case WM_MENUCHAR:        lReturn = pMenu_MenuChar(pMenu, wParam, lParam):        bHandled = True
    Case WM_MENUSELECT:      lReturn = pMenu_MenuSelect(pMenu, wParam, lParam):      bHandled = True
    Case WM_INITMENUPOPUP, _
         WM_UNINITMENUPOPUP: lReturn = pMenu_InitPopup(pMenu, iMsg, wParam, lParam): bHandled = True
        
    Case WM_TIMER:           bHandled = pMenu_Timer(pMenu, wParam)
        
    End Select
End Sub

Friend Property Get fSupportFontPropPage() As pcSupportFontPropPage
    Set fSupportFontPropPage = moFontPage
End Property









'####################################
'##   PUBLIC FUNCTIONS             ##
'####################################

Public Function ContainerKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer) As cPopupMenuItem
    Dim lpMenu As Long: lpMenu = mtMenus.pMenu
    Dim lpItem As Long
    
    Dim liKey As Long:  liKey = KeyCode
    Dim liMask As Long: liMask = Shift
    
    Do While lpMenu
        lpItem = pMenu_ContainerKeyDown(lpMenu, liKey, liMask)
        If lpItem Then Exit Do
        lpMenu = PopupMenu_pMenuNext(lpMenu)
    Loop
    
    If lpItem Then
        Set ContainerKeyDown = pItem(lpItem)
        RaiseEvent Click(ContainerKeyDown)
        KeyCode = ZeroL
        Shift = ZeroL
    End If
    
End Function

Public Function NewMenu() As cPopupMenu
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a new popup menu.
'---------------------------------------------------------------------------------------
    Dim lpMenu As Long
    lpMenu = PopupMenu_Initialize()
    pMenu_Insert lpMenu
    Set NewMenu = pMenu(lpMenu)
End Function

Public Sub SetBackgroundPicture(ByVal oNew As Object)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a picture that is used as the background of the menu.
'---------------------------------------------------------------------------------------
    With mtMenus
        Set .oDcBackground = Nothing
        Set .oDcBackgroundLight = Nothing
        Set .oDcBackgroundSuperLight = Nothing
        
        If Not oNew Is Nothing Then
            Set .oDcBackground = New pcMemDC
            .oDcBackground.CreateFromPicture oNew
        End If
    End With
    
    pMenus_ProcessPicture
End Sub


'####################################
'##   PUBLIC PROPERTIES            ##
'####################################

Public Property Get BackgroundPictureExists() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether a background picture is being used.
'---------------------------------------------------------------------------------------
    BackgroundPictureExists = Not mtMenus.oDcBackground Is Nothing
End Property

Public Property Get ButtonHighlight() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the highlight is drawn with a raised border.
'---------------------------------------------------------------------------------------
    ButtonHighlight = pMenus_DrawStyle(mnuButtonHighlight)
End Property
Public Property Let ButtonHighlight(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the highlight is drawn with a raised border.
'---------------------------------------------------------------------------------------
    pMenus_DrawStyle(mnuButtonHighlight) = bNew
    PropertyChanged PROP_Flags
End Property

Public Property Get ChevronWasClicked() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether infrequently used items are now shown
'             because the user has clicked a chevron.
'---------------------------------------------------------------------------------------
    ChevronWasClicked = mbChevronClicked
End Property

Public Property Get ColorActiveBack() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the backcolor used for active items.
'---------------------------------------------------------------------------------------
    ColorActiveBack = mtMenus.iActiveBackColor
End Property
Public Property Let ColorActiveBack(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the backcolor used for active items.
'---------------------------------------------------------------------------------------
    mtMenus.iActiveBackColor = iNew
    PropertyChanged PROP_ActiveBackColor
End Property

Public Property Get ColorActiveFore() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the forecolor used for active items.
'---------------------------------------------------------------------------------------
    ColorActiveFore = mtMenus.iActiveForeColor
End Property
Public Property Let ColorActiveFore(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the forecolor used for active items.
'---------------------------------------------------------------------------------------
    mtMenus.iActiveForeColor = iNew
    PropertyChanged PROP_ActiveForeColor
End Property

Public Property Get ColorInactiveBack() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the backcolor used for inactive items.
'---------------------------------------------------------------------------------------
    ColorInactiveBack = mtMenus.iInActiveBackColor
End Property
Public Property Let ColorInactiveBack(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the backcolor used for inactive items.
'---------------------------------------------------------------------------------------
    mtMenus.iInActiveBackColor = iNew
    PropertyChanged PROP_InactiveBackColor
End Property

Public Property Get ColorInactiveFore() As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the forecolor used for inactive items.
'---------------------------------------------------------------------------------------
    ColorInactiveFore = mtMenus.iInActiveForeColor
End Property
Public Property Let ColorInactiveFore(ByVal iNew As OLE_COLOR)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the forecolor used for inactive items.
'---------------------------------------------------------------------------------------
    mtMenus.iInActiveForeColor = iNew
    PropertyChanged PROP_InactiveForeColor
End Property

Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the font used by the menus.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property
Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the font used by the menus.
'---------------------------------------------------------------------------------------
    Set moFont = oNew
    If moFont Is Nothing Then Set moFont = Font_CreateDefault(Ambient.Font)
    pMenus_SetFont
    PropertyChanged PROP_Font
End Property

Public Property Get GradientHighlight() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the highlight is drawn with a gradient.
'---------------------------------------------------------------------------------------
    GradientHighlight = pMenus_DrawStyle(mnuGradientHighlight)
End Property
Public Property Let GradientHighlight(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the highlight is drawn with a gradient.
'---------------------------------------------------------------------------------------
    pMenus_DrawStyle(mnuGradientHighlight) = bNew
    PropertyChanged PROP_Flags
End Property

Public Property Get ImageList() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the imagelist used by the menus.
'---------------------------------------------------------------------------------------
    Set ImageList = moImageList
End Property
Public Property Set ImageList(ByVal oNew As cImageList)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set the imagelist used by the menus.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Set moImageList = Nothing
    Set moImageListEvent = Nothing
    Set moImageList = oNew
    Set moImageListEvent = oNew
    On Error GoTo 0
    pMenus_SetImagelist
End Property

Public Property Get ImageProcessBitmap() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the background picture is processed for
'             highlights and infrequenly used items.
'---------------------------------------------------------------------------------------
    ImageProcessBitmap = pMenus_DrawStyle(mnuImageProcessBitmap)
End Property
Public Property Let ImageProcessBitmap(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the background picture is processed for
'             highlights and infrequenly used items.
'---------------------------------------------------------------------------------------
    pMenus_DrawStyle(mnuImageProcessBitmap) = bNew
    PropertyChanged PROP_Flags
End Property

Public Property Get OfficeXPStyle() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the menus are drawn similar to Office XP.
'---------------------------------------------------------------------------------------
    OfficeXPStyle = pMenus_DrawStyle(mnuOfficeXPStyle)
End Property
Public Property Let OfficeXPStyle(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the menus are drawn similar to Office XP.
'---------------------------------------------------------------------------------------
    pMenus_DrawStyle(mnuOfficeXPStyle) = bNew
    PropertyChanged PROP_Flags
End Property

Public Property Get ShowInfrequent() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether infrequently used items are visible.
'---------------------------------------------------------------------------------------
    ShowInfrequent = pMenus_DrawStyle(mnuShowInfrequent)
End Property
Public Property Let ShowInfrequent(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether infrequently used items are visible.
'---------------------------------------------------------------------------------------
    pMenus_DrawStyle(mnuShowInfrequent) = bNew
    mbChevronClicked = False
    PropertyChanged PROP_Flags
End Property

Public Property Get ShowInfrequentHoverDelay() As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating the delay before infrequent items are shown when
'             the mouse hovers over a chevron.  XP only.
'---------------------------------------------------------------------------------------
    ShowInfrequentHoverDelay = mtMenus.iInfreqShowDelay
End Property
Public Property Let ShowInfrequentHoverDelay(ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating the delay before infrequent items are shown when
'             the mouse hovers over a chevron.  XP only.
'---------------------------------------------------------------------------------------
    mtMenus.iInfreqShowDelay = iNew
    PropertyChanged PROP_InfreqShowDelay
End Property

Public Property Get TitleSeparators() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a value indicating whether the separators are drawn as title items.
'---------------------------------------------------------------------------------------
    TitleSeparators = pMenus_DrawStyle(mnuTitleSeparators)
End Property
Public Property Let TitleSeparators(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set a value indicating whether the separators are drawn as title items.
'---------------------------------------------------------------------------------------
    pMenus_DrawStyle(mnuTitleSeparators) = bNew
    PropertyChanged PROP_Flags
End Property

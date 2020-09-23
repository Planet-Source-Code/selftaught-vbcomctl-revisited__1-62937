Attribute VB_Name = "mPopupDraw"
'==================================================================================================
'mPopupDraw.bas                      7/15/05
'
'           PURPOSE:
'               Measure and draw custom popup menu items.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Menus/Popup_Menu_ActiveX_DLL/VB6_PopupMenu_DLL_Full_Source.asp
'               cPopupMenu.cls
'
'==================================================================================================
Option Explicit

Public Type tMenus
    hWndOwner               As Long
    hIml                    As Long
    iIconSize               As Long
    iInfreqShowDelay        As Long
    iItemHeight             As Long
    iFlags                  As Long
    
    hFont                   As Long
    hFontBold               As Long
    
    iBackgroundWidth        As Long
    iBackgroundHeight       As Long
    
    oDcBackground           As pcMemDC
    oDcBackgroundLight      As pcMemDC
    oDcBackgroundSuperLight As pcMemDC
    
    pMenu                   As Long
    
    iActiveForeColor        As OLE_COLOR
    iInActiveForeColor      As OLE_COLOR
    iInActiveBackColor      As OLE_COLOR
    iActiveBackColor        As OLE_COLOR
End Type

Public Const mnuChevron                 As Long = 1024
Public Const mnuShowCheckAndIcon        As Long = 1

'Private Type MEASUREITEMSTRUCT
'    CtlType As Long
'    CtlID As Long
'    itemID As Long
'    itemWidth As Long
'    ItemHeight As Long
'    ItemData As Long
'End Type

Private Const MIS_CtlType As Long = 0
Private Const MIS_ItemWidth As Long = 12
Private Const MIS_ItemHeight As Long = 16
Private Const MIS_ItemData As Long = 20


'Private Type DRAWITEMSTRUCT
'   CtlType As Long
'   CtlID As Long
'   itemID As Long
'   itemAction As Long
'   itemState As Long
'   hwndItem As Long
'   hdc As Long
'   rcItem as RECT
'   ItemData As Long
'End Type

Private Const DIS_CtlType As Long = 0
Private Const DIS_ItemState As Long = 16
Private Const DIS_Hdc As Long = 24
Private Const DIS_rcItem As Long = 28
Private Const DIS_ItemData As Long = 44

'modular because we don't want to allocate all these variables over and over
'again for each item drawn only to have to pass them to all the private procedures.
Private mbRadioCheck                As Boolean
Private mbDisabled                  As Boolean
Private mbChecked                   As Boolean
Private mbHighlighted               As Boolean
Private mbHeader                    As Boolean
Private mbSeparator                 As Boolean
Private mbDefault                   As Boolean
Private mbInfrequent                As Boolean
Private mbNextInfrequent            As Boolean
Private mbPrevInfrequent            As Boolean
Private mbChevron                   As Boolean
Private mbOfficeXPStyle             As Boolean
Private mbShowCheckAndIcon          As Boolean
Private mtRectItem                  As RECT
Private mtRectLeft                  As RECT
Private mtRectRight                 As RECT
Private mtRectIcon1                 As RECT
Private mtRectIcon2                 As RECT
Private mtRectCaption               As RECT
Private mtRectSideBar               As RECT
Private miState                     As Long
Private miFlags                     As Long
Private mhPen                       As Long
Private mhPenOld                    As Long
Private mhFont                      As Long
Private mhFontOld                   As Long
Private mhIml                       As Long
Private miYOffset                   As Long
Private miIconIndex                 As Long
Private miWidth                     As Long
Private mhDc                        As Long
Private miActiveForeColor           As Long
Private miInActiveForeColor         As Long
Private miActiveBackColor           As Long
Private miInActiveBackColor         As Long
Private mtJunk                      As POINT

Private miIconSize                  As Long

Private mpCaption                   As Long
Private mpShortcutDisplay           As Long

Private mpItem                      As Long
Private mpMenu                      As Long

Private mhMenuFont                  As Long
Private mhMenuFontBold              As Long

Private moBackground                As pcMemDC
Private moBackgroundLight           As pcMemDC
Private moBackgroundSuperLight      As pcMemDC

Private miSidebarWidth              As Long
Private miSidebarHeight             As Long
Private mhDcSidebar                 As Long

Private moWorkDC                    As pcMemDC
Private miWorkDCRefCount            As Long
Private miCheckWidth                As Long
Private miRadioCheckWidth           As Long
Private mhFontSymbol                As Long
Private moFont                      As cFont

Public Function PopupMenus_DrawItem(ByRef tMenus As tMenus, ByVal pDIS As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw a popup menu item specified by the DIS structure.
'---------------------------------------------------------------------------------------
    If MemOffset32(pDIS, DIS_CtlType) = ODT_MENU Then
        
        PopupMenus_DrawItem = OneL
        
        mpItem = MemOffset32(pDIS, DIS_ItemData)
        mpMenu = PopupItem_pMenuParent(mpItem)
        
        miState = PopupItem_Style(mpItem)
        miFlags = tMenus.iFlags
        miIconSize = pMeasure_GetIconSize(tMenus)
        
        If Not (CBool(miState And mnuInvisible) Or (CBool(miState And mnuInfrequent) And Not CBool(miFlags And mnuShowInfrequent))) Then
            If tMenus.hIml Then
                mhIml = tMenus.hIml
                miIconIndex = PopupItem_IconIndex(mpItem)
            Else
                miIconIndex = NegOneL
                mhIml = ZeroL
            End If
            
            mbShowCheckAndIcon = CBool(PopupMenu_Style(mpMenu) And mnuShowCheckAndIcon)
            
            Set moBackground = tMenus.oDcBackground
            Set moBackgroundLight = tMenus.oDcBackgroundLight
            Set moBackgroundSuperLight = tMenus.oDcBackgroundSuperLight
            
            mhMenuFont = tMenus.hFont
            mhMenuFontBold = tMenus.hFontBold
            
            miSidebarWidth = PopupMenu_SidebarWidth(mpMenu)
            miSidebarHeight = PopupMenu_SidebarHeight(mpMenu)
            mhDcSidebar = PopupMenu_SidebarHdc(mpMenu)
            
            mpCaption = PopupItem_lpCaption(mpItem)
            mpShortcutDisplay = PopupItem_lpShortcutDisplay(mpItem)
            
            mbRadioCheck = CBool(miState And mnuRadioChecked)
            mbDisabled = CBool(miState And mnuDisabled)
            mbChecked = CBool(miState And mnuChecked) Or mbRadioCheck
            
            mbSeparator = CBool(miState And mnuSeparator)
            mbHeader = mbSeparator And CBool(miFlags And mnuTitleSeparators)
            
            mbDefault = CBool(miState And mnuDefault)
            mbInfrequent = CBool(miState And mnuInfrequent)
            mbChevron = CBool(miState And mnuChevron)
            mbOfficeXPStyle = CBool(miFlags And mnuOfficeXPStyle)
            
            mbHighlighted = CBool(MemOffset32(pDIS, DIS_ItemState) And 1&) And Not (mbDisabled Or mbHeader Or mbSeparator)
            
            With tMenus
                miActiveBackColor = .iActiveBackColor
                miActiveForeColor = .iActiveForeColor
                miInActiveForeColor = .iInActiveForeColor
                miInActiveBackColor = .iInActiveBackColor
            End With
            
            Dim ltItem As RECT
            
            CopyRect ltItem, ByVal UnsignedAdd(pDIS, DIS_rcItem)
            
            With ltItem
                miYOffset = .Top
                mtRectItem.Top = ZeroL
                mtRectItem.Left = ZeroL
                mtRectItem.Right = .Right - .Left
                mtRectItem.bottom = .bottom - .Top
            End With
            
            pDraw_GetRects mtRectItem, mtRectSideBar, mtRectLeft, mtRectRight, mtRectIcon1, mtRectIcon2, mtRectCaption
            
            miWidth = mtRectItem.Right - mtRectItem.Left
            
            mhDc = moWorkDC.hDc
            
            pDraw_Sidebar tMenus
            
            pDraw_Background
            
            If mbChevron Then
                pDraw_Chevron
            ElseIf mbSeparator Then
                pDraw_Separator
            Else
                pDraw_Icon1
                If mbShowCheckAndIcon Then pDraw_Icon mtRectIcon2
                pDraw_Caption
            End If
            
            pDraw_IconBorder
            
            With ltItem
                BitBlt MemOffset32(pDIS, DIS_Hdc), .Left, .Top, .Right - .Left, .bottom - .Top, mhDc, ZeroL, ZeroL, vbSrcCopy
            End With
            
            Set moBackground = Nothing
            Set moBackgroundLight = Nothing
            Set moBackgroundSuperLight = Nothing
            
        End If
    End If
End Function


Private Sub pDraw_Background()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw the background bitmap or color(s).
'---------------------------------------------------------------------------------------
    SetBkMode mhDc, OPAQUE
    Dim liBackColor As Long
    
    If Not mbHighlighted Then
        If mbHeader Then
            If mpCaption _
                Then liBackColor = pDraw_DarkerColor(miInActiveBackColor) _
                Else liBackColor = miInActiveBackColor
        Else
            liBackColor = miInActiveBackColor
            If mbInfrequent Then liBackColor = pDraw_LighterColor(liBackColor)
        End If

        If mbOfficeXPStyle Then
            pDraw_FillRect mtRectRight, liBackColor, IIf(mbInfrequent, moBackgroundLight, moBackground)
            pDraw_FillRect mtRectLeft, pDraw_LighterColor(liBackColor), IIf(mbInfrequent, moBackgroundSuperLight, moBackgroundLight)
        Else
            pDraw_FillRect mtRectItem, liBackColor, IIf(mbInfrequent, moBackgroundLight, moBackground)
        End If

    Else
        Select Case True
        Case mbChevron, CBool(miFlags And mnuButtonHighlight)
            liBackColor = miInActiveBackColor
            If mbInfrequent Then liBackColor = pDraw_LighterColor(liBackColor)
        Case CBool(miFlags And mnuGradientHighlight)
            pDraw_Gradient mhDc, mtRectItem, miInActiveBackColor, miActiveBackColor, False
            liBackColor = NegOneL
        Case mbOfficeXPStyle
            liBackColor = pDraw_LighterColor(pDraw_LighterColor(pDraw_BlendColor(miActiveBackColor, miInActiveBackColor)))
        Case Else
            liBackColor = miActiveBackColor
        End Select

        If liBackColor <> NegOneL Then
            pDraw_FillRect mtRectItem, liBackColor, moBackgroundSuperLight
            If moBackgroundSuperLight Is Nothing And ((mbOfficeXPStyle Or CBool(miFlags And mnuButtonHighlight)) = False) And (miIconIndex > NegOneL Or mbChecked) Then
                Dim liRight As Long
                liRight = mtRectLeft.Right
                If mtRectLeft.Right > mtRectIcon2.Right + 4& Then mtRectLeft.Right = mtRectIcon2.Right + 4&
                If mbInfrequent Then
                    pDraw_FillRect mtRectLeft, pDraw_LighterColor(miInActiveBackColor), Nothing
                Else
                    pDraw_FillRect mtRectLeft, miInActiveBackColor, Nothing
                End If
                mtRectLeft.Right = liRight
            ElseIf moBackgroundSuperLight Is Nothing And mbOfficeXPStyle And CBool(miFlags And mnuButtonHighlight) Then
                pDraw_FillRect mtRectLeft, pDraw_LighterColor(liBackColor), IIf(mbInfrequent, moBackgroundSuperLight, moBackgroundLight)
            End If
        End If

        If CBool(miFlags And mnuButtonHighlight) Or mbChevron Then
            pDraw_Edge mhDc, mtRectItem, EDGE_RAISED, BF_RECT, False
        ElseIf mbOfficeXPStyle Then
            pDraw_Edge mhDc, mtRectItem, 0, 0, True
        End If

    End If

    If Not (mbHighlighted And (CBool(miFlags And mnuOfficeXPStyle) Or (miFlags And mnuButtonHighlight))) And CBool(miFlags And mnuShowInfrequent) Then
        pDraw_GetInfrequentStates
        If (mbInfrequent Xor mbPrevInfrequent) Then
            mhPen = GdiMgr_CreatePen(PS_SOLID, 1&, TranslateColor(IIf(mbPrevInfrequent, vbWhite, &H505050)))
            mhPenOld = SelectObject(mhDc, mhPen)
            MoveToEx mhDc, mtRectItem.Left, mtRectItem.Top, mtJunk
            LineTo mhDc, mtRectItem.Right, mtRectItem.Top
            SelectObject mhDc, mhPenOld
            GdiMgr_DeletePen mhPen
        End If
        If (mbInfrequent Xor mbNextInfrequent) Then
            mhPen = GdiMgr_CreatePen(PS_SOLID, 1&, IIf(mbNextInfrequent, TranslateColor(vb3DShadow), pDraw_BlendColor(TranslateColor(miInActiveBackColor), TranslateColor(vb3DShadow))))
            mhPenOld = SelectObject(mhDc, mhPen)
            MoveToEx mhDc, mtRectItem.Left, mtRectItem.bottom - 1&, mtJunk
            LineTo mhDc, mtRectItem.Right, mtRectItem.bottom - 1&
            SelectObject mhDc, mhPenOld
            GdiMgr_DeletePen mhPen
        End If
    End If
    SetBkMode mhDc, Transparent
End Sub

Private Sub pDraw_Gradient(ByVal hDc As Long, ByRef tR As RECT, ByVal iColorFrom As OLE_COLOR, ByVal iColorTo As OLE_COLOR, ByVal bVertical As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Delegate to the DrawGradient sub with a tRect.
'---------------------------------------------------------------------------------------
    DrawGradient hDc, tR.Left, tR.Top, tR.Right - tR.Left, tR.bottom - tR.Top, iColorFrom, iColorTo, bVertical
End Sub

Private Sub pDraw_Separator()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw the separator or title item.
'---------------------------------------------------------------------------------------
    Dim ltSep As RECT
    
    With mtRectItem
        ltSep.Left = IIf(mbOfficeXPStyle, mtRectRight.Left, .Left) + 2
        ltSep.Right = .Right - 2
        ltSep.Top = ((.bottom - .Top) \ 2) + .Top - IIf(mbOfficeXPStyle, 1, 2)
        ltSep.bottom = ltSep.Top + IIf(mbOfficeXPStyle, -1, 2)
    End With
    
    If mpCaption = ZeroL Then
        OffsetRect ltSep, 0, 1
        pDraw_Edge mhDc, ltSep, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, mbOfficeXPStyle
    Else
        mhFontOld = SelectObject(mhDc, pDraw_GetFontHandle(True))
        
        SetBkMode mhDc, Transparent
        If mbHighlighted And moBackgroundSuperLight Is Nothing And (CBool(miFlags And mnuButtonHighlight) = False) _
                Then SetTextColor mhDc, TranslateColor(miActiveForeColor) _
                Else SetTextColor mhDc, TranslateColor(miInActiveForeColor)
        
        If mbHeader Then
            DrawText mhDc, ByVal mpCaption, -1, mtRectItem, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
        Else
            Dim ltRectCaption As RECT
            Dim liWidth As Long
            
            DrawText mhDc, ByVal mpCaption, -1, ltRectCaption, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
            
            liWidth = ltRectCaption.Right - ltRectCaption.Left
            
            If mbOfficeXPStyle Then
                LSet ltRectCaption = mtRectRight
                ltRectCaption.Right = ltRectCaption.Right - 2
                ltRectCaption.Left = ltRectCaption.Left + 4&
                pDraw_Text mhDc, mpCaption, ltRectCaption, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE, mbDisabled, mbOfficeXPStyle
                ltSep.Left = ltRectCaption.Left + liWidth + 2&
                If ltSep.Left < ltSep.Right Then pDraw_Edge mhDc, ltSep, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, mbOfficeXPStyle
            Else
                LSet ltRectCaption = mtRectItem
                If (mtRectItem.Left + ((mtRectItem.Right - mtRectItem.Left) \ 2) - (liWidth \ 2)) > ltRectCaption.Left Then ltRectCaption.Left = mtRectItem.Left + ((mtRectItem.Right - mtRectItem.Left) \ 2) - (liWidth \ 2)
                ltRectCaption.Right = ltRectCaption.Left + liWidth
                pDraw_Text mhDc, mpCaption, ltRectCaption, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE, mbDisabled, mbOfficeXPStyle
                
                ltSep.Right = ltRectCaption.Left - 2
                pDraw_Edge mhDc, ltSep, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, mbOfficeXPStyle
                ltSep.Right = mtRectItem.Right - 2
                ltSep.Left = ltRectCaption.Right + 2
                pDraw_Edge mhDc, ltSep, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, mbOfficeXPStyle
            End If
        End If
        SelectObject mhDc, mhFontOld
    End If


End Sub

Private Sub pDraw_Chevron()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw a chevron arrow.  Thanks Steve!
'---------------------------------------------------------------------------------------
    Dim ltTemp As RECT
    LSet ltTemp = mtRectItem
    With ltTemp
        .Top = .bottom - 14
        .Right = .Right - 2
        .bottom = .bottom - 1
    End With

    mhPen = GdiMgr_CreatePen(PS_SOLID, 1, TranslateColor(miInActiveForeColor))
    mhPenOld = SelectObject(mhDc, mhPen)

    ltTemp.Left = ((ltTemp.Right - ltTemp.Left) \ 2) - 3 + ltTemp.Left
    ltTemp.Top = ltTemp.Top + 2

    MoveToEx mhDc, ltTemp.Left, ltTemp.Top, mtJunk
    LineTo mhDc, ltTemp.Left + 3, ltTemp.Top + 3
    MoveToEx mhDc, ltTemp.Left, ltTemp.Top + 1, mtJunk
    LineTo mhDc, ltTemp.Left + 3, ltTemp.Top + 3 + 1

    MoveToEx mhDc, ltTemp.Left, ltTemp.Top + 4, mtJunk
    LineTo mhDc, ltTemp.Left + 3, ltTemp.Top + 3 + 4
    MoveToEx mhDc, ltTemp.Left, ltTemp.Top + 1 + 4, mtJunk
    LineTo mhDc, ltTemp.Left + 3, ltTemp.Top + 3 + 1 + 4

    MoveToEx mhDc, ltTemp.Left + 4, ltTemp.Top, mtJunk
    LineTo mhDc, ltTemp.Left + 4 - 3, ltTemp.Top + 3
    MoveToEx mhDc, ltTemp.Left + 4, ltTemp.Top + 1, mtJunk
    LineTo mhDc, ltTemp.Left + 4 - 3, ltTemp.Top + 3 + 1

    MoveToEx mhDc, ltTemp.Left + 4, ltTemp.Top + 4, mtJunk
    LineTo mhDc, ltTemp.Left + 4 - 3, ltTemp.Top + 3 + 4
    MoveToEx mhDc, ltTemp.Left + 4, ltTemp.Top + 1 + 4, mtJunk
    LineTo mhDc, ltTemp.Left + 4 - 3, ltTemp.Top + 3 + 1 + 4

    SelectObject mhDc, mhPenOld
    GdiMgr_DeletePen mhPen

End Sub

Private Sub pDraw_Caption()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw the caption of a menu item.
'---------------------------------------------------------------------------------------

    SetBkMode mhDc, Transparent
    If mbHighlighted And moBackgroundSuperLight Is Nothing And (CBool(miFlags And mnuButtonHighlight) = False) _
            Then SetTextColor mhDc, TranslateColor(miActiveForeColor) _
            Else SetTextColor mhDc, TranslateColor(miInActiveForeColor)

    
    mhFont = pDraw_GetFontHandle(mbDefault)
    
    If mhFont Then
    
        mhFontOld = SelectObject(mhDc, mhFont)
        
        If mhFontOld Then
            
            If mpCaption Then
                pDraw_Text mhDc, mpCaption, mtRectCaption, DT_SINGLELINE Or DT_VCENTER Or DT_LEFT, mbDisabled, mbOfficeXPStyle
            End If
            
            If mpShortcutDisplay Then
                pDraw_Text mhDc, mpShortcutDisplay, mtRectCaption, DT_SINGLELINE Or DT_VCENTER Or DT_RIGHT, mbDisabled, mbOfficeXPStyle
            End If
            SelectObject mhDc, mhFontOld
            
        End If
    End If
End Sub

Private Sub pDraw_Icon1()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw an icon or a checkbox in the leftmost position.
'---------------------------------------------------------------------------------------
    If mbChecked Then
        pDraw_Edge mhDc, mtRectIcon1, BDR_SUNKENOUTER, BF_RECT, mbOfficeXPStyle
        mhFontOld = SelectObject(mhDc, mhFontSymbol)
        SetTextColor mhDc, miInActiveForeColor

        Dim liWidth As Long
        liWidth = IIf(mbRadioCheck, miRadioCheckWidth, miCheckWidth)
        With mtRectIcon1
            .Left = .Left + (.Right - .Left) \ TwoL - (liWidth \ TwoL)
            If Not mbRadioCheck Or (.bottom - .Top) > 15 Then .Top = .Top + 1
        End With
        
        pDraw_Text mhDc, StrPtr(StrConv(IIf(mbRadioCheck, "h", "b") & vbNullChar, vbFromUnicode)), mtRectIcon1, DT_SINGLELINE Or DT_VCENTER, mbDisabled, False

        SelectObject mhDc, mhFontOld

    ElseIf Not mbShowCheckAndIcon Then
        pDraw_Icon mtRectIcon1

    End If

End Sub

Private Sub pDraw_IconBorder()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw a raised border around an icon or checkbox to show that it is highlighted.
'---------------------------------------------------------------------------------------
    If mbHighlighted And ((mbChecked Or mbRadioCheck) Or (miIconIndex > NegOneL And mhIml <> ZeroL)) And Not (mbOfficeXPStyle Or mbDisabled Or mbChevron Or CBool(miFlags And (mnuGradientHighlight))) Then
        Dim ltTemp As RECT
        If Not CBool(miFlags And mnuButtonHighlight) Then
            LSet ltTemp = mtRectLeft
            InflateRect ltTemp, -TwoL, -TwoL
            If mbInfrequent = mbPrevInfrequent Then ltTemp.Top = ltTemp.Top - OneL
            If mbInfrequent = mbNextInfrequent Then ltTemp.bottom = ltTemp.bottom + OneL
            DrawEdge mhDc, ltTemp, BDR_RAISEDINNER, BF_RECT
        End If
    End If
End Sub

Private Sub pDraw_Sidebar(ByRef tMenus As tMenus)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw one item's worth of a sidebar picture along the left edge.
'---------------------------------------------------------------------------------------
    If mtRectSideBar.Right > mtRectSideBar.Left Then
        
        If mhDcSidebar Then
            Dim liYOffset As Long
            Dim lpItem As Long
            
            lpItem = mpItem
            
            Do
                liYOffset = liYOffset + pMeasure_GetItemHeight(tMenus, lpItem)
                lpItem = PopupItem_pItemNext(lpItem)
            Loop While lpItem
            
            With mtRectSideBar
                BitBlt mhDc, .Left, .Top, .Right - .Left, .bottom - .Top, mhDcSidebar, 0, miSidebarHeight - liYOffset, vbSrcCopy
            End With
        End If
    End If
End Sub

Private Sub pDraw_Icon(ByRef tRect As RECT)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw the item icon in the appropriate state.
'---------------------------------------------------------------------------------------
   If miIconIndex > NegOneL Then
        If mbDisabled Then
            pDraw_ImageListIconDisabled mhDc, mhIml, miIconIndex, tRect.Left, tRect.Top, miIconSize
        Else
            If mbHighlighted And mbOfficeXPStyle Then
                pDraw_ImageListIconDisabled mhDc, mhIml, miIconIndex, tRect.Left + 1, tRect.Top + 1, miIconSize, True
                pDraw_ImageListIcon mhDc, mhIml, miIconIndex, tRect.Left - 1, tRect.Top - 1
            Else
                pDraw_ImageListIcon mhDc, mhIml, miIconIndex, tRect.Left, tRect.Top
            End If
        End If
    End If
End Sub

Private Sub pDraw_GetRects(ByRef tRectItem As RECT, ByRef tSideBar As RECT, ByRef tLeft As RECT, ByRef tRight As RECT, ByRef tIcon1 As RECT, ByRef tIcon2 As RECT, ByRef tCaption As RECT)
    'IN: tRectItem
    'OUT: tRectItem and all other rects
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : get the rects of various parts of the menu item.
'---------------------------------------------------------------------------------------
    Const SPACING As Long = 4

    Dim liNumIcons As Long
    Dim liWidth As Long
    Dim liHeight As Long

    'store the width and height that the rectangles are to fill.
    liWidth = tRectItem.Right - tRectItem.Left
    liHeight = tRectItem.bottom - tRectItem.Top

    'result is 1 or 2 icons.
    liNumIcons = 1& + Abs(mbShowCheckAndIcon)

    'Side bar is always on the left
    LSet tSideBar = tRectItem
    'if there is a picture to put on the sidebar then add the width
    If mhDcSidebar Then
        tSideBar.Right = tSideBar.Left + miSidebarWidth
        liWidth = liWidth + miSidebarWidth
    Else
        tSideBar.Right = tSideBar.Left
    End If

    'the item (background, icons, caption) will be drawn to the right of the sidebar
    tRectItem.Left = tSideBar.Right

    'start out with the left and right sections equal to the whole item
    LSet tLeft = tRectItem
    LSet tRight = tRectItem

    'the left item extends only to cover the icons
    tLeft.Right = tLeft.Left + SPACING + (miIconSize + SPACING) * liNumIcons

    'the right item takes everything else
    tRight.Left = tLeft.Right

    'the left icon will be a square with sides = liiconsize, and will be centered
    'in the available height.  the left edge will be SPACING pixels to the right of the whole item
    With tIcon1
        .Top = tLeft.Top + ((tLeft.bottom - tLeft.Top) \ 2&) - (miIconSize \ 2&)
        .bottom = .Top + miIconSize
        .Left = SPACING + tLeft.Left
        .Right = .Left + miIconSize
    End With
    
    'if check and icon
    If liNumIcons = 2& Then
        'the next icon will also be centered vertically, but will be SPACING pixels to the right of the
        'first icon
        With tIcon2
            .Top = tLeft.Top + ((tLeft.bottom - tLeft.Top) \ 2&) - (miIconSize \ 2&)
            .bottom = .Top + miIconSize
            .Left = SPACING + tIcon1.Right
            .Right = .Left + miIconSize
        End With
    Else
        'no space for this icon
        tIcon2.Left = tIcon1.Right
        tIcon2.Right = tIcon1.Right
    End If

    'the caption will begin at the right item
    LSet tCaption = tRectItem
    'indent the text eight pixels
    tCaption.Left = tRight.Left + SPACING
    'shortcut captions are right aligned with a larger space from the right edge
    tCaption.Right = tCaption.Right - SPACING - SPACING
    
End Sub

Private Function pDraw_GetFontHandle(ByVal bBold As Boolean) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a font handle for text operations.
'---------------------------------------------------------------------------------------
    If bBold _
        Then pDraw_GetFontHandle = mhMenuFontBold _
        Else pDraw_GetFontHandle = mhMenuFont
End Function

Private Sub pDraw_GetInfrequentStates()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Get the infrequent states of the previous and next items for drawing
'             a border between frequent/infrequent items.
'---------------------------------------------------------------------------------------
    Dim lpItem As Long
    
    lpItem = PopupItem_pItemNext(mpItem)
    Do While lpItem
        If Not CBool(PopupItem_Style(lpItem) And mnuInvisible) Then Exit Do
        lpItem = PopupItem_pItemNext(lpItem)
    Loop
    
    If lpItem _
        Then mbNextInfrequent = CBool(PopupItem_Style(lpItem) And mnuInfrequent) _
        Else mbNextInfrequent = mbInfrequent
    
    lpItem = PopupItem_pItemPrev(mpItem)
    Do While lpItem
        If Not CBool(PopupItem_Style(lpItem) And mnuInvisible) Then Exit Do
        lpItem = PopupItem_pItemPrev(lpItem)
    Loop
    
    If lpItem _
        Then mbPrevInfrequent = CBool(PopupItem_Style(lpItem) And mnuInfrequent) _
        Else mbPrevInfrequent = mbInfrequent

End Sub

Private Sub pDraw_FillRect(ByRef tR As RECT, ByVal iColor As Long, ByVal oDc As pcMemDC)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Fill a rect with a picture if it is available, otherwise with a backcolor.
'---------------------------------------------------------------------------------------
    If oDc Is Nothing Then
        Dim hBr As Long
        hBr = GdiMgr_CreateSolidBrush(TranslateColor(iColor))
        If hBr Then
            FillRect mhDc, tR, hBr
            GdiMgr_DeleteBrush hBr
        End If
    Else
        PopupMenus_TileArea mhDc, tR.Left, tR.Top, tR.Right - tR.Left, tR.bottom - tR.Top, _
                 oDc.hDc, oDc.Width, oDc.Height, miYOffset
    End If
End Sub

Private Function pDraw_Text(ByVal lhDc As Long, ByVal lpText As Long, tR As RECT, ByVal dtFlags As Long, ByVal bDisabled As Boolean, ByVal bOfficeXP As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw text in an enabled or disabled state.
'---------------------------------------------------------------------------------------
   If bDisabled Then
      If bOfficeXP Then
         SetTextColor lhDc, TranslateColor(vb3DShadow)
      Else
         SetTextColor lhDc, TranslateColor(vb3DHighlight)
         OffsetRect tR, OneL, OneL
      End If
   End If
   DrawText lhDc, ByVal lpText, NegOneL, tR, dtFlags
   If bDisabled Then
      If Not bOfficeXP Then
         OffsetRect tR, NegOneL, NegOneL
         SetTextColor lhDc, TranslateColor(vbButtonShadow)
         DrawText lhDc, ByVal lpText, NegOneL, tR, dtFlags
      End If
   End If
End Function

Private Sub pDraw_Edge( _
      ByVal hDc As Long, _
      ByRef qrc As RECT, _
      ByVal edge As Long, _
      ByVal grfFlags As Long, _
      ByVal bOfficeXPStyle As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw an edge around a given rectangle.
'---------------------------------------------------------------------------------------
    If (bOfficeXPStyle) Then
        Dim junk As POINT
        Dim hPenOld As Long
        Dim hPen As Long
        If (qrc.bottom > qrc.Top) Then
            hPen = GdiMgr_CreatePen(PS_SOLID, 1, TranslateColor(vbHighlight))
        Else
            hPen = GdiMgr_CreatePen(PS_SOLID, 1, TranslateColor(vb3DShadow))
        End If
        hPenOld = SelectObject(hDc, hPen)
        MoveToEx hDc, qrc.Left, qrc.Top, junk
        LineTo hDc, qrc.Right - 1, qrc.Top
        If (qrc.bottom > qrc.Top) Then
            LineTo hDc, qrc.Right - 1, qrc.bottom - 1
            LineTo hDc, qrc.Left, qrc.bottom - 1
            LineTo hDc, qrc.Left, qrc.Top
        End If
        SelectObject hDc, hPenOld
        GdiMgr_DeletePen hPen
    Else
        DrawEdge hDc, qrc, edge, grfFlags
    End If
End Sub

Private Sub pDraw_ImageListIcon( _
        ByVal hDc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        Optional ByVal bSelected As Boolean = False, _
        Optional ByVal bBlend25 As Boolean = False _
    )
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw an icon from an imagelist.
'---------------------------------------------------------------------------------------
    ImageList_Draw hIml, iIconIndex, hDc, lX, lY, _
                   ILD_TRANSPARENT Or (-bSelected * ILD_SELECTED) Or (-bBlend25 * ILD_BLEND25)
End Sub

Private Sub pDraw_ImageListIconDisabled( _
        ByVal hDc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lSize As Long, _
        Optional ByVal bDisabled As Boolean _
    )
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Draw an grayed icon from an imagelist.
'---------------------------------------------------------------------------------------

    Dim hIcon As Long

    hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)

    If hIcon Then
        If bDisabled Then
            DrawState hDc, GetSysColorBrush(vb3DShadow And &H1F), 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO
        Else
            DrawState hDc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED
        End If
        DestroyIcon hIcon
    End If

End Sub

Private Function pDraw_BlendColor(ByVal iColorFrom As OLE_COLOR, ByVal iColorTo As OLE_COLOR) As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Split the difference between two colors.
'---------------------------------------------------------------------------------------

    iColorFrom = TranslateColor(iColorFrom)
    iColorTo = TranslateColor(iColorTo)

    Dim R As Long
    Dim G As Long
    Dim b As Long

    R = (iColorFrom And &HFF) + ((iColorTo And &HFF) - (iColorFrom And &HFF)) \ 2
    If (R > 255) Then R = 255 Else If (R < 0) Then R = 0
    G = ((iColorFrom \ &H100) And &HFF&) + (((iColorTo \ &H100) And &HFF&) - ((iColorFrom \ &H100) And &HFF&)) \ 2
    If (G > 255) Then G = 255 Else If (G < 0) Then G = 0
    b = ((iColorFrom \ &H10000) And &HFF&) + (((iColorTo \ &H10000) And &HFF&) - ((iColorFrom \ &H10000) And &HFF&)) \ 2
    If (b > 255) Then b = 255 Else If (b < 0) Then b = 0

    pDraw_BlendColor = RGB(R, G, b)
End Function

Private Function pDraw_LighterColor(ByVal iColor As OLE_COLOR) As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a similar, lighter color.
'---------------------------------------------------------------------------------------

Dim h As Single, s As Single, L As Single
Dim R As Long, G As Long, b As Long

    iColor = TranslateColor(iColor)

    R = iColor And &HFF&
    G = (iColor \ &H100) And &HFF&
    b = (iColor \ &H10000) And &HFF&

    PopupMenus_HLSforRGB R, G, b, h, L, s

    If (L > 0.99) Then
        L = L * 0.8
    Else
        L = L * 1.1
        If (L > 1) Then L = 1
    End If

    PopupMenus_RGBforHLS h, L, s, R, G, b

    pDraw_LighterColor = RGB(R, G, b)

End Function

Private Function pDraw_DarkerColor(ByVal iColor As OLE_COLOR) As OLE_COLOR
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return a similar, darker color.
'---------------------------------------------------------------------------------------

Dim h As Single, s As Single, L As Single
Dim R As Long, G As Long, b As Long

    iColor = TranslateColor(iColor)

    R = iColor And &HFF&
    G = (iColor \ &H100) And &HFF&
    b = (iColor \ &H10000) And &HFF&


    PopupMenus_HLSforRGB R, G, b, h, L, s

    L = L - 0.15
    If L < 0 Then L = 0

    PopupMenus_RGBforHLS h, L, s, R, G, b

    pDraw_DarkerColor = RGB(R, G, b)

End Function








Public Function PopupMenus_MeasureItem(ByRef tMenus As tMenus, ByVal lpMIS As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Measure the item specified by the MIS structure.
'---------------------------------------------------------------------------------------
    If MemOffset32(lpMIS, MIS_CtlType) = ODT_MENU Then
        PopupMenus_MeasureItem = OneL
        
        Dim liWidth As Long
        Dim lhDc As Long
        Dim liWidthCaption As Long
        Dim liWidthShortcut As Long
        Dim hFontOld As Long
        Dim liStyle As Long
        
        Dim lpItem As Long: lpItem = MemOffset32(lpMIS, MIS_ItemData)
        Dim lpMenu As Long: lpMenu = PopupItem_pMenuParent(lpItem)
        
        liStyle = PopupItem_Style(lpItem)
        
        lhDc = moWorkDC.hDc
    
        If lhDc Then
            hFontOld = SelectObject(lhDc, IIf(CBool(liStyle And (mnuDefault Or mnuSeparator)), tMenus.hFontBold, tMenus.hFont))
            If hFontOld Then
                pMeasure_GetTextExtent lhDc, PopupItem_lpCaption(lpItem), liWidthCaption
                pMeasure_GetTextExtent lhDc, PopupItem_lpShortcutDisplay(lpItem), liWidthShortcut
                SelectObject lhDc, hFontOld
            End If
        End If
    
        If PopupMenu_SidebarHdc(lpMenu) Then liWidth = liWidth + PopupMenu_SidebarWidth(lpMenu)
    
        liWidth = liWidth + liWidthCaption + liWidthShortcut + 10& + ((pMeasure_GetIconSize(tMenus) + 4&) * (-CBool(PopupMenu_Style(lpMenu) And mnuShowCheckAndIcon) + OneL))
        
        MemOffset32(lpMIS, MIS_ItemWidth) = liWidth
        MemOffset32(lpMIS, MIS_ItemHeight) = pMeasure_GetItemHeight(tMenus, lpItem)
    End If
End Function

Private Function pMeasure_GetItemHeight(ByRef tMenus As tMenus, ByVal pItem As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Calculate the height of the given item.
'---------------------------------------------------------------------------------------
    Dim liStyle As Long
    liStyle = PopupItem_Style(pItem)
    
    If CBool(liStyle And mnuInvisible) Or (CBool(liStyle And mnuInfrequent) And Not CBool(tMenus.iFlags And mnuShowInfrequent)) Then
        pMeasure_GetItemHeight = ZeroL
    ElseIf (liStyle And mnuSeparator) Then
        If PopupItem_lpCaption(pItem) = ZeroL Then
            If CBool(tMenus.iFlags And mnuOfficeXPStyle) _
                Then pMeasure_GetItemHeight = 3& _
                Else pMeasure_GetItemHeight = 8&
        Else
            pMeasure_GetItemHeight = tMenus.iItemHeight - 4&
            If pMeasure_GetItemHeight < 8& Then pMeasure_GetItemHeight = 8&
        End If
        
    ElseIf liStyle And mnuChevron Then
        pMeasure_GetItemHeight = 16&
        
    Else
        pMeasure_GetItemHeight = tMenus.iItemHeight
        
    End If
End Function

Private Sub pMeasure_GetTextExtent(ByVal lhDc As Long, ByVal lpString As Long, Optional ByRef iWidth As Long, Optional ByRef iHeight As Long)
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the height and width of the string in the hdc.
'---------------------------------------------------------------------------------------
    Dim tP As SIZE

    iHeight = ZeroL
    iWidth = ZeroL

    If lpString Then
        If GetTextExtentPoint32(lhDc, ByVal lpString, lstrlen(lpString), tP) Then
            iWidth = tP.cx
            iHeight = tP.cy
        End If
    End If

End Sub

Private Function pMeasure_GetIconSize(ByRef tMenus As tMenus) As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Return the size (width and height are equal) of an icon.
'---------------------------------------------------------------------------------------
    pMeasure_GetIconSize = tMenus.iIconSize
    If pMeasure_GetIconSize <= ZeroL Then pMeasure_GetIconSize = GetSystemMetrics(SM_CXMENUCHECK) + OneL
    If pMeasure_GetIconSize > tMenus.iItemHeight - 4& Then pMeasure_GetIconSize = tMenus.iItemHeight - 4&
End Function


Public Function PopupMenus_GetDC() As Long
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Set up modular state for drawing operations.  Return a shared dc handle.
'---------------------------------------------------------------------------------------
    If miWorkDCRefCount = ZeroL Then
        Set moWorkDC = New pcMemDC
        moWorkDC.Create Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY \ 4&
        Set moFont = New cFont
        moFont.FaceName = "Marlett"
        moFont.Height = -15&
        moFont.Charset = fntCharsetSymbol
        mhFontSymbol = moFont.GetHandle()
        miRadioCheckWidth = moFont.TextWidth("h", moWorkDC.hDc) + 1
        miCheckWidth = moFont.TextWidth("b", moWorkDC.hDc)
    End If
    miWorkDCRefCount = miWorkDCRefCount + OneL
    PopupMenus_GetDC = moWorkDC.hDc
End Function

Public Sub PopupMenus_ReleaseDC()
'---------------------------------------------------------------------------------------
' Date      : 7/15/05
' Purpose   : Decrement the reference count and destroy the work dc if no longer needed.
'---------------------------------------------------------------------------------------
    miWorkDCRefCount = miWorkDCRefCount - OneL
    If miWorkDCRefCount = ZeroL Then
        moFont.ReleaseHandle mhFontSymbol
        mhFontSymbol = ZeroL
        Set moFont = Nothing
        Set moWorkDC = Nothing
    End If
End Sub


Public Sub PopupMenus_TileArea( _
            ByVal hDcDst As Long, _
            ByVal xDst As Long, _
            ByVal yDst As Long, _
            ByVal cxDst As Long, _
            ByVal cyDst As Long, _
            ByVal hDcSrc As Long, _
            ByVal cxSrc As Long, _
            ByVal cySrc As Long, _
   Optional ByVal cyOffset As Long, _
   Optional ByVal cxOffset As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Repeat a source dc image vertically and horizontally on a destination dc.
' Lineage   : www.vbaccelerator.com
'---------------------------------------------------------------------------------------

Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((xDst + cxOffset) Mod cxSrc)
    lSrcStartY = ((yDst + cyOffset) Mod cySrc)
    lSrcStartWidth = (cxSrc - lSrcStartX)
    lSrcStartHeight = (cySrc - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY

    lDstY = yDst
    lDstHeight = lSrcStartHeight

    Do While lDstY < (yDst + cyDst)
        If (lDstY + lDstHeight) > (yDst + cyDst) Then
            lDstHeight = yDst + cyDst - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = xDst
        lSrcX = lSrcStartX
        Do While lDstX < (xDst + cxDst)
            If (lDstX + lDstWidth) > (xDst + cxDst) Then
                lDstWidth = xDst + cxDst - lDstX
                If (lDstWidth = ZeroL) Then lDstWidth = 4&
            End If
            BitBlt hDcDst, lDstX, lDstY, lDstWidth, lDstHeight, hDcSrc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = cxSrc
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = cySrc
    Loop
End Sub

Public Sub PopupMenus_HLSforRGB( _
            ByVal R As Long, _
            ByVal G As Long, _
            ByVal b As Long, _
            ByRef h As Single, _
            ByRef L As Single, _
            ByRef s As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Translate Color values
' Lineage   : www.vbaccelerator.com
'---------------------------------------------------------------------------------------
            
 Dim Max As Single
 Dim Min As Single
 Dim delta As Single
 Dim rR As Single, rG As Single, rB As Single

     rR = R / 255: rG = G / 255: rB = b / 255

 '{Given: rgb each in [0,1].
 ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
         Max = pMax(rR, rG, rB)
         Min = pMin(rR, rG, rB)
             L = (Max + Min) / 2 '{This is the lightness}
         '{Next calculate saturation}
         If Max = Min Then
             'begin {Acrhomatic case}
             s = 0
             h = 0
             'end {Acrhomatic case}
         Else
             'begin {Chromatic case}
                 '{First calculate the saturation.}
             delta = Max - Min
             If L <= 0.5 Then
                 s = delta / (Max + Min)
             Else
                 s = delta / (2 - Max - Min)
             End If
             '{Next calculate the hue.}
             
             If rR = Max Then
                 h = (rG - rB) / delta     '{Resulting color is between yellow and magenta}
             ElseIf rG = Max Then
                 h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
             ElseIf rB = Max Then
                 h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
             End If
         'end {Chromatic Case}
     End If
End Sub

Public Sub PopupMenus_RGBforHLS( _
            ByVal h As Single, _
            ByVal L As Single, _
            ByVal s As Single, _
            ByRef R As Long, _
            ByRef G As Long, _
            ByRef b As Long)

'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Translate Color values
' Lineage   : www.vbaccelerator.com
'---------------------------------------------------------------------------------------
            
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single

    If s = 0 Then
    ' Achromatic case:
    rR = L: rG = L: rB = L
    Else
    ' Chromatic case:
    ' delta = Max-Min
    If L <= 0.5 Then
        's = (Max - Min) / (Max + Min)
        ' Get Min value:
        Min = L * (1 - s)
    Else
        's = (Max - Min) / (2 - Max - Min)
        ' Get Min value:
        Min = L - s * (1 - L)
    End If
    ' Get the Max value:
    Max = 2 * L - Min
    
    ' Now depending on sector we can evaluate the h,l,s:
    If (h < 1) Then
        rR = Max
        If (h < 0) Then
            rG = Min
            rB = rG - h * (Max - Min)
        Else
            rB = Min
            rG = h * (Max - Min) + rB
        End If
    ElseIf (h < 3) Then
        rG = Max
        If (h < 2) Then
            rB = Min
            rR = rB - (h - 2) * (Max - Min)
        Else
            rR = Min
            rB = (h - 2) * (Max - Min) + rR
        End If
    Else
        rB = Max
        If (h < 4) Then
            rR = Min
            rG = rR - (h - 4) * (Max - Min)
        Else
            rG = Min
            rR = (h - 4) * (Max - Min) + rG
        End If
        
    End If
            
    End If
    R = rR * 255: G = rG * 255: b = rB * 255
End Sub

Private Function pMax(rR As Single, rG As Single, rB As Single) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Translate Color values
' Lineage   : www.vbaccelerator.com
'---------------------------------------------------------------------------------------
    If (rR > rG) Then
        If (rR > rB) Then
            pMax = rR
        Else
            pMax = rB
        End If
    Else
        If (rB > rG) Then
            pMax = rB
        Else
            pMax = rG
        End If
    End If
End Function
Private Function pMin(rR As Single, rG As Single, rB As Single) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/21/05
' Purpose   : Translate Color values
' Lineage   : www.vbaccelerator.com
'---------------------------------------------------------------------------------------
    If (rR < rG) Then
        If (rR < rB) Then
            pMin = rR
        Else
            pMin = rB
        End If
    Else
        If (rB < rG) Then
            pMin = rB
        Else
            pMin = rG
        End If
    End If
End Function


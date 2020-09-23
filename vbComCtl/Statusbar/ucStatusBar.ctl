VERSION 5.00
Begin VB.UserControl ucStatusBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   HasDC           =   0   'False
   PropertyPages   =   "ucStatusBar.ctx":0000
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ToolboxBitmap   =   "ucStatusBar.ctx":000D
End
Attribute VB_Name = "ucStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================================
'ucStatusBar.ctl        12/15/04
'
'           PURPOSE:
'               Implement the comctl32 statusbar control.
'
'           LINEAGE:
'               http://www.vbaccelerator.com/home/VB/Code/Controls/Status_Bar/Status_Bar_Control/VB6_Status_Bar_Control_Full_Source.asp
'               vbalSbar.ctl
'
'==================================================================================================

Option Explicit

Public Enum eStatusBarPanelBorder
    sbarInset = &H0&
    sbarNoBorders = SBT_NOBORDERS
    sbarRaised = SBT_POPOUT
End Enum

Public Enum eStatusBarPanelStyle
    sbarStandard = ZeroL
    sbarCaps = &H10000000 Or SBT_OWNERDRAW
    sbarNum = &H20000000 Or SBT_OWNERDRAW
    sbarIns = &H30000000 Or SBT_OWNERDRAW
    sbarScrl = &H40000000 Or SBT_OWNERDRAW
    sbarTime = &H50000000 Or SBT_OWNERDRAW
    sbarDate = &H60000000 Or SBT_OWNERDRAW
    sbarDateTime = &H70000000 Or SBT_OWNERDRAW
End Enum

Public Event PanelClick(ByVal oPanel As cPanel, ByVal iButton As evbComCtlMouseButton)
Public Event PanelDblClick(ByVal oPanel As cPanel, ByVal iButton As evbComCtlMouseButton)

Implements iSubclass

Private Const PROP_Font = "Font"
Private Const PROP_SizeGrip = "SizeGrip"
Private Const PROP_SimpleMode = "SimpleMode"
Private Const PROP_SimpleBorder = "SimpleBorder"
Private Const PROP_BorderStyle = "BorderStyle"
Private Const PROP_Themeable = "Themeable"

Private Const DEF_SizeGrip = True
Private Const DEF_SimpleMode = False
Private Const DEF_SimpleBorder = ZeroL
Private Const DEF_BorderStyle = vbccBorderNone
Private Const DEF_Themeable = True

Private Type tStatusPanel
    iId As Long
    iIconIndex As Long
    hIcon As Long
    sText As String
    sKey As String
    sToolTipText As String
    iMinWidth As Long
    iIdealWidth As Long
    iStyle As eStatusBarPanelStyle
    bSpring As Boolean
    bFit As Boolean
    bEnabled As Boolean
End Type

Private mtPanels()                  As tStatusPanel
Private miPanelsCount               As Long
Private miPanelsUbound              As Long

Private miPanelsControl             As Long

Private mhWnd                       As Long
Private WithEvents moFont           As cFont
Attribute moFont.VB_VarHelpID = -1
Private WithEvents moFontPage       As pcSupportFontPropPage
Attribute moFontPage.VB_VarHelpID = -1
Private WithEvents moImageListEvent As cImageList
Attribute moImageListEvent.VB_VarHelpID = -1
Private moImageList                 As cImageList

Private mbSizeGrip                  As Boolean
Private mbSimpleMode                As Boolean
Private miSimpleBorder              As eStatusBarPanelBorder

Private mbThemeable                 As Boolean

Private mhFont                      As Long
Private miIconSize                  As Long

Private msSimpleText                As String

Private miBorderStyle               As evbComCtlBorderStyle

Private Const ucStatusBar = "StatusBar"
Private Const cPanels = "cPanels"
Private Const cPanel = "cPanel"

Private Const NMHDR_hwndFrom        As Long = 0
Private Const NMHDR_code            As Long = 8

'Private Const NMMOUSE_ItemSpec      As Long = 12

Private Const NMMOUSE_pt_x          As Long = 20
'Private Const NMMOUSE_pt_y          As Long = 24

Private Const DRAWITEMSTRUCT_ItemId As Long = 8
Private Const DRAWITEMSTRUCT_hdc    As Long = 24
Private Const DRAWITEMSTRUCT_rcItem As Long = 28

Private Const TIMERID               As Long = 2
Private Const TIMERFREQ             As Long = 350

Private Const BORDERMASK            As Long = sbarInset Or sbarNoBorders Or sbarRaised
Private Const STYLEMASK             As Long = sbarStandard Or sbarCaps Or sbarNum Or sbarIns Or sbarScrl Or sbarTime Or sbarDate Or sbarDateTime

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long): End Sub
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Respond to notifications from the status bar.
'---------------------------------------------------------------------------------------
    If uMsg = WM_NOTIFY Then
        If MemOffset32(lParam, NMHDR_hwndFrom) = mhWnd Then
            bHandled = True
            Select Case MemOffset32(lParam, NMHDR_code)
            Case NM_CLICK, NM_RCLICK, NM_DBLCLK, NM_RDBLCLK
                pClick MemOffset32(lParam, NMHDR_code), MemOffset32(lParam, NMMOUSE_pt_x)
            End Select
        End If
    ElseIf uMsg = WM_TIMER Then
        bHandled = True
        pUpdateCustomItems
    ElseIf uMsg = WM_DRAWITEM Then
        bHandled = True
        pDrawItem MemOffset32(lParam, DRAWITEMSTRUCT_ItemId), MemOffset32(lParam, DRAWITEMSTRUCT_hdc), UnsignedAdd(lParam, DRAWITEMSTRUCT_rcItem)
    End If
End Sub

Private Sub moImageListEvent_Changed()
    Set ImageList = moImageListEvent
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If StrComp(PropertyName, "Font") = ZeroL Then moFont.OnAmbientFontChanged Ambient.Font
End Sub

Private Sub UserControl_Initialize()
    'On Error Resume Next
    LoadShellMod
    InitCC ICC_BAR_CLASSES
    msSimpleText = vbNullChar
    Set moFontPage = New pcSupportFontPropPage
    'On Error GoTo 0
End Sub

Private Sub UserControl_InitProperties()
    'On Error Resume Next
    Set moFont = Font_CreateDefault(Ambient.Font)
    mbSizeGrip = DEF_SizeGrip
    mbSimpleMode = DEF_SimpleMode
    miSimpleBorder = DEF_SimpleBorder
    miBorderStyle = DEF_BorderStyle
    mbThemeable = DEF_Themeable
    pCreate
    'On Error GoTo 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'On Error Resume Next
    Set moFont = Font_Read(PropBag, PROP_Font, Ambient.Font)
    mbSizeGrip = PropBag.ReadProperty(PROP_SizeGrip, DEF_SizeGrip)
    mbSimpleMode = PropBag.ReadProperty(PROP_SimpleMode, DEF_SimpleMode)
    miSimpleBorder = PropBag.ReadProperty(PROP_SimpleBorder, DEF_SimpleBorder)
    miBorderStyle = PropBag.ReadProperty(PROP_BorderStyle, DEF_BorderStyle)
    mbThemeable = PropBag.ReadProperty(PROP_Themeable, DEF_Themeable)
    pCreate
    'On Error GoTo 0
End Sub

Private Sub UserControl_Resize()
    'On Error Resume Next
    'If Ambient.UserMode = False Then
    Static bInHere As Boolean
    If Not bInHere Then
        bInHere = True
        pSize
        pSetParts
        bInHere = False
    End If
    'End If
    'On Error GoTo 0
End Sub

Private Sub UserControl_Terminate()
    'On Error Resume Next
    fPanels_Clear
    pDestroy
    If mhFont Then moFont.ReleaseHandle mhFont
    ReleaseShellMod
    Set moFontPage = Nothing
    'On Error GoTo 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'On Error Resume Next
    Font_Write moFont, PropBag, PROP_Font
    PropBag.WriteProperty PROP_SizeGrip, mbSizeGrip, DEF_SizeGrip
    PropBag.WriteProperty PROP_SimpleMode, mbSimpleMode, DEF_SimpleMode
    PropBag.WriteProperty PROP_SimpleBorder, miSimpleBorder, DEF_SimpleBorder
    PropBag.WriteProperty PROP_BorderStyle, miBorderStyle, DEF_BorderStyle
    PropBag.WriteProperty PROP_Themeable, mbThemeable, DEF_Themeable
    'On Error GoTo 0
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

Private Sub pCreate()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Create the status bar and install the subclass.
'---------------------------------------------------------------------------------------
    pDestroy
    
    Dim lsAnsi As String
    lsAnsi = StrConv(WC_STATUSBAR & vbNullChar, vbFromUnicode)
    
    mhWnd = CreateWindowEx(ZeroL, StrPtr(lsAnsi), ZeroL, WS_CHILD Or WS_VISIBLE Or SBT_TOOLTIPS Or (-mbSizeGrip * SBARS_SIZEGRIP), ZeroL, ZeroL, ScaleWidth, ScaleHeight, UserControl.hWnd, ZeroL, App.hInstance, ByVal ZeroL)
    
    If mhWnd Then
        If Ambient.UserMode Then
            Subclass_Install Me, UserControl.hWnd, Array(WM_NOTIFY, WM_TIMER, WM_DRAWITEM)
        End If
        
        EnableWindowTheme mhWnd, mbThemeable
        
        pSetFont
        pSetParts
        pSetBorder
        pSize
        pSetPanels
        pCheckTimer
        
        pPanel_SetText SB_SIMPLEID
        SendMessage mhWnd, SB_SIMPLE, -mbSimpleMode, ZeroL
        
    End If

End Sub

Private Sub pDestroy()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Destroy the status bar window and the subclass.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        KillTimer UserControl.hWnd, TIMERID
        Subclass_Remove Me, UserControl.hWnd
        DestroyWindow mhWnd
        mhWnd = ZeroL
    End If
End Sub

Private Function pItem(ByVal iIndex As Long) As cPanel
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a cPanel object for the given index.
'---------------------------------------------------------------------------------------
    Debug.Assert iIndex > NegOneL And iIndex < miPanelsCount
    
    If iIndex > NegOneL And iIndex < miPanelsCount Then
        Set pItem = New cPanel
        pItem.fInit Me, mtPanels(iIndex).iId, iIndex
    End If
End Function

Private Sub pClick(ByVal iMsg As Long, ByVal x As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Figure out which panel was clicked, and raise the event.
'---------------------------------------------------------------------------------------
    
    If Not (miPanelsCount = ZeroL Or mhWnd = ZeroL Or mbSimpleMode) Then
        Dim liParts() As Long
        Dim liBorders(0 To 2) As Long
        Dim i As Long
        
        ReDim liParts(ZeroL To miPanelsCount - OneL)
        
        SendMessage mhWnd, SB_GETPARTS, miPanelsCount, VarPtr(liParts(ZeroL))
        SendMessage mhWnd, SB_GETBORDERS, ZeroL, VarPtr(liBorders(0))
        
        For i = ZeroL To miPanelsCount - OneL
            If liParts(i) < OneL Or liParts(i) > x Then Exit For
        Next
        
        If i > ZeroL Then
            If x < (liParts(i - OneL) + liBorders(2)) Then i = miPanelsCount
        End If
        
        If i < miPanelsCount Then
            Select Case iMsg
            Case NM_CLICK
                RaiseEvent PanelClick(pItem(i), vbccMouseLButton)
            Case NM_RCLICK
                RaiseEvent PanelClick(pItem(i), vbccMouseRButton)
            Case NM_DBLCLK
                RaiseEvent PanelDblClick(pItem(i), vbccMouseLButton)
            Case NM_RDBLCLK
                RaiseEvent PanelDblClick(pItem(i), vbccMouseRButton)
            End Select
        End If
        
    End If
End Sub

Private Sub pPropChanged(ByRef s As String)
    If Ambient.UserMode = False Then PropertyChanged s
End Sub

Private Sub pUpdateCustomItems()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return update date/time/keyboard state panels.
'---------------------------------------------------------------------------------------
    Dim i As Long
    Dim tR As RECT
    Dim bUpdate As Boolean
    Dim ls As String
    Dim b As Boolean
    Dim bNewString As Boolean
    
    If mhWnd Then
        For i = ZeroL To miPanelsCount - OneL
            If CBool(mtPanels(i).iStyle And Not SBT_MASK) Then
                pGetCustomItem i, ls, b
                If CBool(StrComp(ls, mtPanels(i).sText)) Then
                    pSize
                    InvalidateRect mhWnd, ByVal ZeroL, OneL
                    bUpdate = True
                    Exit For
                ElseIf (b Xor mtPanels(i).bEnabled) Then
                    If SendMessage(mhWnd, SB_GETRECT, i, VarPtr(tR)) _
                        Then InvalidateRect mhWnd, tR, OneL _
                        Else InvalidateRect mhWnd, ByVal ZeroL, OneL
                    bUpdate = True
                End If
            End If
        Next
        If bUpdate Then UpdateWindow mhWnd
    End If
End Sub

Private Sub pSetParts()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Tell the statusbar what size to make each panel.
'---------------------------------------------------------------------------------------
    Dim liBorders(0 To 2) As Long
    Const BORDERHORZ As Long = ZeroL
    'Const BORDERVERT As Long = OneL
    Const BORDERSPACING As Long = TwoL
    Const GRIPPERSIZE As Long = 16
    
    Dim liParts() As Long
    Dim i As Long
    Dim liTotalWidth As Long
    Dim liWidth As Long
    
    If mhWnd Then
        
        If miPanelsCount > ZeroL Then
        
            SendMessage mhWnd, SB_GETBORDERS, ZeroL, VarPtr(liBorders(0))
            
            ReDim liParts(ZeroL To miPanelsCount - OneL)
            
            For i = ZeroL To miPanelsCount - OneL
                liParts(i) = pGetGoodWidth(i)
                liTotalWidth = liTotalWidth + liParts(i) + liBorders(BORDERSPACING)
            Next
            
            liTotalWidth = liTotalWidth + liBorders(BORDERHORZ) + liBorders(BORDERHORZ)
            If mbSizeGrip Then liTotalWidth = liTotalWidth + GRIPPERSIZE
            
            If liTotalWidth > (ScaleWidth - OneL) Then
                For i = miPanelsCount - OneL To ZeroL Step NegOneL
                    With mtPanels(i)
                        If .bFit = False Then
                            liTotalWidth = liTotalWidth - .iIdealWidth + .iMinWidth
                            liParts(i) = .iMinWidth
                            If liTotalWidth <= (ScaleWidth - OneL) Then Exit For
                        End If
                    End With
                Next
            End If
            
            If liTotalWidth < (ScaleWidth - OneL) Then
                For i = miPanelsCount - OneL To ZeroL Step NegOneL
                    If mtPanels(i).bSpring Then
                        liParts(i) = liParts(i) + ((ScaleWidth - OneL) - liTotalWidth)
                        Exit For
                    End If
                Next
            End If
            
            liTotalWidth = liBorders(BORDERHORZ)
            For i = ZeroL To miPanelsCount - OneL
                liWidth = liParts(i)
                liParts(i) = liTotalWidth + liWidth
                liTotalWidth = liTotalWidth + liWidth + liBorders(BORDERSPACING)
            Next
        
            SendMessage mhWnd, SB_SETPARTS, miPanelsCount, VarPtr(liParts(ZeroL))
            
        Else
            
            SendMessage mhWnd, SB_SETPARTS, ZeroL, ZeroL
            
        End If
        
    End If
End Sub

Private Function pGetGoodWidth(ByVal iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a default width for the given panel.
'---------------------------------------------------------------------------------------
    Debug.Assert iIndex > NegOneL And iIndex < miPanelsCount
    pGetGoodWidth = mtPanels(iIndex).iIdealWidth
    If mhWnd Then
        If mtPanels(iIndex).bFit Then
            pGetGoodWidth = pTextWidth(mtPanels(iIndex).sText)
            If mtPanels(iIndex).iIconIndex > NegOneL Then pGetGoodWidth = pGetGoodWidth + miIconSize
        End If
    End If
End Function

Private Function pTextWidth(ByRef s As String) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the width of the text using the current font.
'---------------------------------------------------------------------------------------
    Dim tSize As SIZE
    Dim lDc As Long
    lDc = GetDC(mhWnd)
    If lDc Then
        Dim lhFontOld As Long
        
        If mhFont Then
            lhFontOld = SelectObject(lDc, mhFont)
            If lhFontOld Then
                If GetTextExtentPoint32(lDc, ByVal StrPtr(s), LenB(s), tSize) Then
                    pTextWidth = tSize.cx + 8&
                End If
                SelectObject lDc, lhFontOld
            End If
        End If
        ReleaseDC mhWnd, lDc
    Else
        Debug.Assert False
        
    End If
End Function

Private Sub pDrawItem(ByVal iIndex As Long, ByVal hDc As Long, ByVal lpRect As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Draw a date/time/keyboard state panel.
'---------------------------------------------------------------------------------------
    If iIndex <> SB_SIMPLEID Then
        Dim bEnabled As Boolean
        Dim sText As String
        Dim tR As RECT
        Dim tSize As SIZE
        Dim liOldBk As Long
        
        CopyMemory tR, ByVal lpRect, 16&
        
        With mtPanels(iIndex)
            bEnabled = .bEnabled
            pGetCustomItem iIndex, sText, bEnabled
            .sText = sText
            .bEnabled = bEnabled
            liOldBk = SetBkMode(hDc, OneL)
            GetTextExtentPoint32 hDc, ByVal StrPtr(sText), LenB(sText) - OneL, tSize
            tR.Top = tR.Top + ((tR.bottom - tR.Top) \ TwoL) - (tSize.cy \ TwoL)
            tR.Right = tR.Right + IIf(.iIconIndex > NegOneL, 4& + miIconSize, TwoL)
            tR.Left = tR.Left + IIf(.iIconIndex > NegOneL, 4& + miIconSize, TwoL)
            DrawState hDc, ZeroL, ZeroL, StrPtr(sText), LenB(sText) - OneL, tR.Left, tR.Top, tR.Right - tR.Left, tR.bottom - tR.Top, DST_TEXT Or ((bEnabled + OneL) * DSS_DISABLED)
            SetBkMode hDc, liOldBk
        End With

    Else
        Debug.Assert False
        
    End If
End Sub

Private Sub pSize()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Calculate the size of the statusbar and the size the usercontrol.
'---------------------------------------------------------------------------------------
    Dim liBorders(0 To 2) As Long
    
    If mhWnd Then
        SendMessage mhWnd, SB_GETBORDERS, ZeroL, VarPtr(liBorders(0))
        liBorders(1) = liBorders(1) + liBorders(1)
        If miBorderStyle <> vbccBorderNone Then liBorders(1) = liBorders(1) + liBorders(1)
        Dim lhDc As Long
        Dim lhFontOld As Long
        Dim ltSize As SIZE
        Dim liHeight As Long
        
        liHeight = ScaleHeight - liBorders(1)
        
        If mhFont Then
            lhDc = GetDC(mhWnd)
            If lhDc Then
                lhFontOld = SelectObject(lhDc, mhFont)
                If lhFontOld Then
                    GetTextExtentPoint32W lhDc, "A", OneL, ltSize
                    liHeight = ltSize.cy + liBorders(1)
                    SelectObject lhDc, lhFontOld
                End If
                ReleaseDC mhWnd, lhDc
            End If
        End If
        
        SendMessage mhWnd, SB_SETMINHEIGHT, liHeight, ZeroL
        SendMessage mhWnd, WM_SIZE, ZeroL, ZeroL
        
        Dim ltRect As RECT
        GetWindowRect mhWnd, ltRect
        UserControl.SIZE ScaleX(ltRect.Right - ltRect.Left, vbPixels, vbTwips), ScaleY(ltRect.bottom - ltRect.Top, vbPixels, vbTwips)
        
    End If
End Sub


Private Sub pSetBorder()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the window border and redraw the statusbar.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SetWindowStyle mhWnd, ZeroL, WS_BORDER
        SetWindowStyleEx mhWnd, ZeroL, WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
        Select Case miBorderStyle
        Case vbccBorderSingle
            SetWindowStyle mhWnd, WS_BORDER, ZeroL
        Case vbccBorderThin
            SetWindowStyleEx mhWnd, WS_EX_STATICEDGE, ZeroL
        Case vbccBorderSunken
            SetWindowStyleEx mhWnd, WS_EX_CLIENTEDGE, ZeroL
        End Select
        SetWindowPos mhWnd, ZeroL, ZeroL, ZeroL, ZeroL, ZeroL, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
        InvalidateRect mhWnd, ByVal ZeroL, OneL
        UpdateWindow mhWnd
    End If
End Sub

Private Sub pGetCustomItem( _
      ByVal iIndex As Long, _
      ByRef sText As String, _
      ByRef bEnabled As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the text and enabled state for a date/time/keyboard state panel.
'---------------------------------------------------------------------------------------

    Const CustomMask As Long = sbarTime Or sbarScrl Or sbarNum Or sbarIns Or sbarDateTime Or sbarDate Or sbarCaps
    Static b(0 To 255) As Byte
    
    bEnabled = True
    sText = vbNullString
    Select Case (mtPanels(iIndex).iStyle And CustomMask)
    Case sbarTime
        sText = Format$(Now, "short time")
    Case sbarScrl
        sText = "SCRL"
        GetKeyState VK_SCROLL
        GetKeyboardState b(0)
        bEnabled = CBool(b(VK_SCROLL))
    Case sbarNum
        sText = "NUM"
        GetKeyState VK_NUMLOCK
        GetKeyboardState b(0)
        bEnabled = CBool(b(VK_NUMLOCK))
    Case sbarIns
        sText = "OVR"
        GetKeyState VK_INSERT
        GetKeyboardState b(0)
        bEnabled = CBool(b(VK_INSERT))
    Case sbarDateTime
        sText = Format$(Date, "M/d/YY") & " " & Format$(Time, "h:mm AMPM")
    Case sbarDate
        sText = Format$(Date, "M/d/YY")
    Case sbarCaps
        sText = "CAPS"
        GetKeyState VK_CAPITAL
        GetKeyboardState b(0)
        bEnabled = CBool(b(VK_CAPITAL))
    End Select
    sText = StrConv(sText & vbNullChar, vbFromUnicode)
   
End Sub

Private Sub pSetPanels()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set all panel information to the statusbar.
'---------------------------------------------------------------------------------------
    pSize
    Dim i As Long
    For i = ZeroL To miPanelsCount - OneL
        pPanel_SetText i
        pPanel_SetToolTip i
        pPanel_SetIcon i
    Next
End Sub

Private Sub pCheckTimer()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Ensure that we only have a timer if there are custom items to update.
'---------------------------------------------------------------------------------------
    Dim i As Long
    Dim bTimer As Boolean
    
    For i = ZeroL To miPanelsCount - OneL
        If CBool(mtPanels(i).iStyle And Not SBT_MASK) Then
            bTimer = True
            Exit For
        End If
    Next
    If mhWnd Then
        KillTimer UserControl.hWnd, TIMERID
        If bTimer Then SetTimer UserControl.hWnd, TIMERID, TIMERFREQ, ZeroL
    End If
    
End Sub

Private Sub pSetFont()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Update the font used by the statusbar.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Dim hFont As Long
    hFont = moFont.GetHandle()
    SendMessage mhWnd, WM_SETFONT, hFont, OneL
    If mhFont Then moFont.ReleaseHandle mhFont
    mhFont = hFont
    pSetParts
    pSetPanels
    pSize
    On Error GoTo 0
End Sub

Private Sub moFont_Changed()
    moFont.OnAmbientFontChanged Ambient.Font
    pSetFont
    pPropChanged PROP_Font
End Sub


Friend Sub fPanels_Enum_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the next cPanel object in an enumeration.
'---------------------------------------------------------------------------------------
    If tEnum.iControl <> miPanelsControl Then gErr vbccCollectionChangedDuringEnum, cPanels
    tEnum.iIndex = tEnum.iIndex + OneL
    bNoMoreItems = (tEnum.iIndex >= miPanelsCount)
    If Not bNoMoreItems Then
        Dim loPanel As cPanel
        Set loPanel = New cPanel
        loPanel.fInit Me, mtPanels(tEnum.iIndex).iId, tEnum.iIndex
        Set vNextItem = loPanel
    End If
End Sub

Friend Sub fPanels_Enum_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Skip a number of panels in an enumeration.
'---------------------------------------------------------------------------------------
    If tEnum.iControl <> miPanelsControl Then gErr vbccCollectionChangedDuringEnum, cPanels
    tEnum.iIndex = tEnum.iIndex + iSkipCount
    bSkippedAll = tEnum.iIndex <= miPanelsCount
End Sub

Friend Property Get fPanels_Control() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return an identity number for the current cPanels collection.
'---------------------------------------------------------------------------------------
    fPanels_Control = miPanelsControl
End Property


Friend Function fPanels_Add( _
        ByRef sText As String, _
        ByRef sKey As String, _
        ByRef sToolTipText As String, _
        ByVal iStyle As eStatusBarPanelStyle, _
        ByVal iBorder As eStatusBarPanelBorder, _
        ByVal iIconIndex As Long, _
        ByVal fMinWidth As Single, _
        ByVal fIdealWidth As Single, _
        ByVal bSpring As Boolean, _
        ByVal bFit As Boolean, _
        ByRef vPanelInsertBefore As Variant) _
            As cPanel
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Add a panel to the collection.
'---------------------------------------------------------------------------------------
    Dim liIndex As Long
    If IsMissing(vPanelInsertBefore) Then
        liIndex = miPanelsCount
    Else
        liIndex = pPanels_GetIndex(vPanelInsertBefore)
        If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, ucStatusBar
    End If
    
    Dim liNewUbound As Long
    liNewUbound = RoundToInterval(miPanelsCount)
    If liNewUbound > miPanelsUbound Then
        miPanelsUbound = liNewUbound
        ReDim Preserve mtPanels(0 To miPanelsUbound)
    End If
    
    If liIndex < miPanelsCount Then
        With mtPanels(miPanelsCount)
            'clear the string references that will be overwritten
            .sKey = vbNullString
            .sText = vbNullString
            .sToolTipText = vbNullString
            
            CopyMemory mtPanels(liIndex + OneL).iId, mtPanels(liIndex).iId, LenB(mtPanels(0)) * (miPanelsCount - liIndex)
            
            'clear the duplicate string references
            ZeroMemory mtPanels(liIndex).iId, LenB(mtPanels(0))
        End With
    End If
    
    miPanelsCount = miPanelsCount + OneL
    
    With mtPanels(liIndex)
        .iStyle = iStyle Or iBorder
        .sText = StrConv(sText & vbNullChar, vbFromUnicode)
        .sToolTipText = StrConv(sToolTipText & vbNullChar, vbFromUnicode)
        .sKey = sKey
        .iIconIndex = iIconIndex
        .iMinWidth = ScaleX(fMinWidth, vbContainerSize, vbPixels)
        .iIdealWidth = ScaleX(fIdealWidth, vbContainerSize, vbPixels)
        If .iMinWidth < 0 Then .iMinWidth = 0
        If .iMinWidth > .iIdealWidth Then .iIdealWidth = .iMinWidth
        .bSpring = bSpring
        .bFit = bFit
        .bEnabled = True
        .iId = NextItemId()
        .hIcon = ZeroL
        If iIconIndex > NegOneL And Not moImageList Is Nothing Then
            .hIcon = ImageList_GetIcon(moImageList.hIml, iIconIndex, ZeroL)
        End If
        
        If CBool(.iStyle And Not SBT_MASK) Then
            pGetCustomItem liIndex, .sText, .bEnabled
        End If
        
    End With
    
    pSetParts
    pSetPanels
    pSize
    pCheckTimer
    
    Incr fPanels_Control
End Function

Friend Sub fPanels_Remove(ByRef vPanel As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove a panel from the collection.
'---------------------------------------------------------------------------------------
    Dim liIndex As Long
    liIndex = pPanels_GetIndex(vPanel)
    If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, ucStatusBar
    
    miPanelsCount = miPanelsCount - OneL
    If liIndex < miPanelsCount Then
        With mtPanels(liIndex)
            If .hIcon Then
                DestroyIcon .hIcon
                .hIcon = ZeroL
            End If
            .sKey = vbNullString
            .sText = vbNullString
            .sToolTipText = vbNullString
            CopyMemory .iId, mtPanels(liIndex + OneL).iId, LenB(mtPanels(0)) * (miPanelsCount - liIndex)
        End With
        ZeroMemory mtPanels(miPanelsCount).iId, LenB(mtPanels(0))
    End If
    
    pSetParts
    pSetPanels
    pSize
    pCheckTimer
    
    Incr fPanels_Control
End Sub

Friend Sub fPanels_Clear()
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Remove all panels from the collection.
'---------------------------------------------------------------------------------------
    Dim i As Long
    For i = ZeroL To miPanelsCount - OneL
        If mtPanels(i).hIcon Then
            DestroyIcon mtPanels(i).hIcon
            mtPanels(i).hIcon = ZeroL
        End If
    Next
    
    pSetParts
    pSize
    
    miPanelsCount = ZeroL
    pCheckTimer
    
    Incr fPanels_Control
End Sub

Friend Function fPanels_Item(ByRef vPanel As Variant) As cPanel
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return an object identifying the given panel.
'---------------------------------------------------------------------------------------
    Dim liIndex As Long
    liIndex = pPanels_GetIndex(vPanel)
    If liIndex = NegOneL Then gErr vbccKeyOrIndexNotFound, ucStatusBar
        
    Set fPanels_Item = New cPanel
    fPanels_Item.fInit Me, mtPanels(liIndex).iId, liIndex
        
End Function

Friend Function fPanels_Exists(ByRef vPanel As Variant)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating if the given panel exists.
'---------------------------------------------------------------------------------------
    fPanels_Exists = pPanels_GetIndex(vPanel) <> NegOneL
End Function

Friend Property Get fPanels_Count() As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the number of panels in the control.
'---------------------------------------------------------------------------------------
    fPanels_Count = miPanelsCount
End Property


Private Function pPanels_GetIndex(ByRef vPanel As Variant) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return an index of a panel given its key, object or index.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    If VarType(vPanel) = vbObject Then
        If TypeOf vPanel Is cPanel Then
            Dim loPanel As cPanel
            Set loPanel = vPanel
            pPanels_GetIndex = NegOneL
            pPanels_GetIndex = loPanel.IconIndex - OneL
        End If
    ElseIf VarType(vPanel) = vbString Then
        pPanels_GetIndex = pPanels_FindString(CStr(vPanel))
    Else
        pPanels_GetIndex = NegOneL
        pPanels_GetIndex = CLng(vPanel) - OneL
        If pPanels_GetIndex < NegOneL Or pPanels_GetIndex >= miPanelsCount Then pPanels_GetIndex = NegOneL
    End If
    On Error GoTo 0
End Function

Private Function pPanels_FindString(ByRef s As String) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Find a panel by key.
'---------------------------------------------------------------------------------------
    If LenB(s) Then
        For pPanels_FindString = ZeroL To miPanelsCount - OneL
            If StrComp(s, mtPanels(pPanels_FindString).sKey) = ZeroL Then Exit Function
        Next
    End If
    pPanels_FindString = NegOneL
End Function



Friend Property Get fPanel_IconIndex(ByVal iId As Long, ByRef iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the iconindex for the given panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_IconIndex = mtPanels(iIndex).iIconIndex
    End If
End Property
Friend Property Let fPanel_IconIndex(ByVal iId As Long, ByRef iIndex As Long, ByVal iNew As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the iconindex for the given panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        With mtPanels(iIndex)
            If .hIcon Then DestroyIcon .hIcon
            .hIcon = ZeroL
            .iIconIndex = iNew
            If Not moImageList Is Nothing Then
                .hIcon = ImageList_GetIcon(moImageList.hIml, iNew, ZeroL)
            End If
        End With
        pPanel_SetIcon iIndex
    End If
End Property
    
Friend Property Get fPanel_Key(ByVal iId As Long, ByRef iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the panel's key.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_Key = mtPanels(iIndex).sKey
    End If
End Property
Friend Property Let fPanel_Key(ByVal iId As Long, ByRef iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the panel's key.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        If pPanels_FindString(sNew) <> NegOneL Then gErr vbccKeyAlreadyExists, cPanel
        mtPanels(iIndex).sKey = sNew
    End If
End Property

Friend Property Get fPanel_ToolTipText(ByVal iId As Long, ByRef iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the panel's tooltiptext.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_ToolTipText = StrConv(LeftB$(mtPanels(iIndex).sToolTipText, LenB(mtPanels(iIndex).sToolTipText) - OneL), vbUnicode)
    End If
End Property
Friend Property Let fPanel_ToolTipText(ByVal iId As Long, ByRef iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the panel's tooltiptext.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        mtPanels(iIndex).sToolTipText = StrConv(sNew & vbNullChar, vbFromUnicode)
        pPanel_SetToolTip iIndex
    End If
End Property

Friend Property Get fPanel_Text(ByVal iId As Long, ByRef iIndex As Long) As String
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the panel's text.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_Text = StrConv(LeftB$(mtPanels(iIndex).sText, LenB(mtPanels(iIndex).sText) - OneL), vbUnicode)
    End If
End Property
Friend Property Let fPanel_Text(ByVal iId As Long, ByRef iIndex As Long, ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the panel's text.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        mtPanels(iIndex).sText = StrConv(sNew & vbNullChar, vbFromUnicode)
        pPanel_SetText iIndex
        If mtPanels(iIndex).bFit Then pSize
    End If
End Property

Friend Property Get fPanel_MinWidth(ByVal iId As Long, ByRef iIndex As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the minimum width for a panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_MinWidth = mtPanels(iIndex).iMinWidth
    End If
End Property

Friend Property Let fPanel_MinWidth(ByVal iId As Long, ByRef iIndex As Long, ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the minimum width for a panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        With mtPanels(iIndex)
            .iMinWidth = ScaleX(fNew, vbContainerSize, vbPixels)
            If .iMinWidth < 0 Then .iMinWidth = 0
            If .iIdealWidth < .iMinWidth Then .iIdealWidth = .iMinWidth
        End With
        pSize
    End If
End Property

Friend Property Get fPanel_IdealWidth(ByVal iId As Long, ByRef iIndex As Long) As Single
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the Ideal width for a panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_IdealWidth = mtPanels(iIndex).iIdealWidth
    End If
End Property
Friend Property Let fPanel_IdealWidth(ByVal iId As Long, ByRef iIndex As Long, ByVal fNew As Single)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the Ideal width for a panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        With mtPanels(iIndex)
            .iIdealWidth = ScaleX(fNew, vbContainerSize, vbPixels)
            If .iIdealWidth < 0 Then .iIdealWidth = 0
            If .iMinWidth > .iIdealWidth Then .iMinWidth = .iIdealWidth
        End With
        pSize
    End If
End Property

Friend Property Get fPanel_Style(ByVal iId As Long, ByRef iIndex As Long) As eStatusBarPanelStyle
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the style of a panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_Style = mtPanels(iIndex).iStyle And STYLEMASK
    End If
End Property
Friend Property Let fPanel_Style(ByVal iId As Long, ByRef iIndex As Long, ByVal iNew As eStatusBarPanelStyle)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the style of a panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        With mtPanels(iIndex)
            .iStyle = (.iStyle And Not STYLEMASK) Or (iNew And STYLEMASK)
        End With
        pPanel_SetText iIndex
        pSize
    End If
End Property

Friend Property Get fPanel_Border(ByVal iId As Long, ByRef iIndex As Long) As eStatusBarPanelBorder
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get the border style of a panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_Border = mtPanels(iIndex).iStyle And BORDERMASK
    End If
End Property
Friend Property Let fPanel_Border(ByVal iId As Long, ByRef iIndex As Long, ByVal iNew As eStatusBarPanelBorder)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the border style of a panel.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        With mtPanels(iIndex)
            .iStyle = (.iStyle And Not BORDERMASK) Or (iNew And BORDERMASK)
        End With
        pPanel_SetText iIndex
        pSize
    End If
End Property

Friend Property Get fPanel_Spring(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get a value indicating whether the panel expands to fill all extra space.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_Spring = mtPanels(iIndex).bSpring
    End If
End Property
Friend Property Let fPanel_Spring(ByVal iId As Long, ByRef iIndex As Long, ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the panel expands to fill all extra space.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        mtPanels(iIndex).bSpring = bNew
        pSize
    End If
End Property

Friend Property Get fPanel_Index(ByVal iId As Long, ByRef iIndex As Long) As Long
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the one-based index of a panel in the collection.
'---------------------------------------------------------------------------------------
    If pPanel_Verify(iId, iIndex) Then
        fPanel_Index = iIndex + OneL
    End If
End Property



Private Function pPanel_Verify(ByVal iId As Long, ByRef iIndex As Long) As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Verify that a panel still exists in the collection.
'---------------------------------------------------------------------------------------
    If iIndex > NegOneL And iIndex < miPanelsCount Then
        pPanel_Verify = CBool(mtPanels(iIndex).iId = iId)
    Else
        iIndex = NegOneL
    End If
    
    If pPanel_Verify = False Then
        For iIndex = ZeroL To miPanelsCount - OneL
            pPanel_Verify = (mtPanels(iIndex).iId = iId)
            If pPanel_Verify Then Exit For
        Next
    End If
    
End Function

Private Sub pPanel_SetText(ByVal iIndex As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the text of a panel.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        If iIndex = SB_SIMPLEID Then
            SendMessage mhWnd, SB_SETTEXT, iIndex Or (miSimpleBorder And SBT_MASK), StrPtr(msSimpleText)
        Else
            SendMessage mhWnd, SB_SETTEXT, iIndex Or (mtPanels(iIndex).iStyle And SBT_MASK), StrPtr(mtPanels(iIndex).sText)
        End If
    End If
End Sub

Private Sub pPanel_SetToolTip(ByVal iIndex As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the tooltiptext of a panel.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, SB_SETTIPTEXT, iIndex, StrPtr(mtPanels(iIndex).sToolTipText)
    End If
End Sub

Private Sub pPanel_SetIcon(ByVal iIndex As Long)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the icon of a panel.
'---------------------------------------------------------------------------------------
    If mhWnd Then
        SendMessage mhWnd, SB_SETICON, iIndex, mtPanels(iIndex).hIcon
    End If
End Sub



Public Property Get Font() As cFont
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the font used by this control.
'---------------------------------------------------------------------------------------
    Set Font = moFont
End Property
Public Property Set Font(ByVal oNew As cFont)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the font used by this control.
'---------------------------------------------------------------------------------------
    If oNew Is Nothing _
        Then Set moFont = Font_CreateDefault(Ambient.Font) _
        Else Set moFont = oNew
    pSetFont
    pPropChanged PROP_Font
End Property

Public Property Get Panels() As cPanels
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a collection of panels.
'---------------------------------------------------------------------------------------
    Set Panels = New cPanels
    Panels.fInit Me
End Property

Public Property Get Simple() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return a value indicating whether simple (one text panel) mode is enabled.
'---------------------------------------------------------------------------------------
    Simple = mbSimpleMode
End Property
Public Property Let Simple(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether simple (one text panel) mode is enabled.
'---------------------------------------------------------------------------------------
    mbSimpleMode = bNew
    If mhWnd Then SendMessage mhWnd, SB_SIMPLE, -mbSimpleMode, ZeroL
    pPropChanged PROP_SimpleMode
End Property

Public Property Get SimpleText() As String
Attribute SimpleText.VB_MemberFlags = "400"
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the string used for the simple mode panel.
'---------------------------------------------------------------------------------------
    SimpleText = StrConv(LeftB$(msSimpleText, LenB(msSimpleText) - OneL), vbUnicode)
End Property
Public Property Let SimpleText(ByRef sNew As String)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the string used for the simple mode panel.
'---------------------------------------------------------------------------------------
    msSimpleText = StrConv(sNew & vbNullChar, vbFromUnicode)
    pPanel_SetText SB_SIMPLEID
End Property

Public Property Get ImageList() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the imagelist used by the control.
'---------------------------------------------------------------------------------------
    Set ImageList = moImageList
End Property
Public Property Set ImageList(ByVal oNew As cImageList)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the imagelist used by this control and update the hIcons.
'---------------------------------------------------------------------------------------
    On Error Resume Next
    Set moImageList = Nothing
    Set moImageListEvent = Nothing
    Set moImageList = oNew
    Set moImageListEvent = oNew
    On Error GoTo 0
    If Not moImageList Is Nothing Then
        If ImageList_GetIconSize(moImageList.hIml, miIconSize, ZeroL) = ZeroL Then miIconSize = ZeroL
    End If
    Dim i As Long
    For i = ZeroL To miPanelsCount - OneL
        With mtPanels(i)
            If .hIcon Then
                DestroyIcon .hIcon
                .hIcon = ZeroL
                .iIconIndex = NegOneL
            End If
        End With
    Next
End Property

Public Property Get BorderStyle() As evbComCtlBorderStyle
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the window's borderstyle.
'---------------------------------------------------------------------------------------
    BorderStyle = miBorderStyle
End Property
Public Property Let BorderStyle(ByVal iNew As evbComCtlBorderStyle)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the window's borderstyle.
'---------------------------------------------------------------------------------------
    miBorderStyle = iNew
    pSetBorder
    pPropChanged PROP_BorderStyle
End Property

Public Property Get SizeGrip() As Boolean
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Get whether the size grip is enabled.  This only works as expected for
'             statusbars that are aligned to the bottom of the form.
'---------------------------------------------------------------------------------------
    SizeGrip = mbSizeGrip
End Property
Public Property Let SizeGrip(ByVal bNew As Boolean)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set whether the size grip is enabled.  This only works as expected for
'             statusbars that are aligned to the bottom of the form.
'---------------------------------------------------------------------------------------
    If mbSizeGrip Xor bNew Then
        mbSizeGrip = bNew
        pCreate
        pPropChanged PROP_SizeGrip
    End If
End Property

Public Property Get SimpleBorder() As eStatusBarPanelBorder
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Return the border style used for the simple mode panel.
'---------------------------------------------------------------------------------------
    SimpleBorder = miSimpleBorder
End Property
Public Property Let SimpleBorder(ByVal iNew As eStatusBarPanelBorder)
'---------------------------------------------------------------------------------------
' Date      : 2/26/05
' Purpose   : Set the border style used for the simple mode panel.
'---------------------------------------------------------------------------------------
    miSimpleBorder = iNew And BORDERMASK
    pPanel_SetText SB_SIMPLEID
    pPropChanged PROP_SimpleBorder
    If mbSimpleMode Then
        If mhWnd Then
            InvalidateRect mhWnd, ByVal ZeroL, OneL
            UpdateWindow mhWnd
        End If
    End If
End Property

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
        If mhWnd Then EnableWindowTheme mhWnd, mbThemeable
    End If
End Property

